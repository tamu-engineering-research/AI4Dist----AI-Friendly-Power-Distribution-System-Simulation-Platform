import win32com.client
import networkx as nx
import matplotlib.pyplot as plt
import csv, os
import numpy as np
import pandas as pd
import random
from .dss_tools import *
from rl.core import Env



class env(Env):
    def __init__(self, case_path, agents, market=None, params=None):
        # unpack configuration dic
        self.ts = params['time_step']
        self.maxStep = params['max_step']
        self.case = dssCase(case_path, self.ts)
        self.caseName = case_path.split('\\')[-1].split('.')[0]

        self.loadProfile = None
        self.pvProfile = None
        self.windProfile = None
        self.profileIndex = None
        self.hourNum = None
        self.DEREnable = params['DEREnable']
        
        # load load and DER profiles if supplied
        if not params['load_profile'] == None:
            self.loadProfile = pd.read_csv(params['load_profile'], index_col='date_time', parse_dates=True)

        if not params['pv_profile'] == None:
            self.pvProfile = pd.read_csv(params['pv_profile'], index_col='date_time', parse_dates=True)

        if not params['wind_profile'] == None:
            self.windProfile = pd.read_csv(params['wind_profile'], index_col='date_time', parse_dates=True)

        # required fields for gym
        self.svNum = None
        self.actNum = None
        self.acton_space = None
        self.observation_space = None

        # market
        self.market = market

        # properties of agents
        self.agents = agents
        self.agentNum = len(self.agents)
        self.trainingAgentIdx = 0
        self.marketAgentIdx = []
        self.ssAgentIdx = []
        self.dynamicAgentIdx = []
        for a in range(self.agentNum):
            if self.agents[a].timescale == "market":
                self.marketAgentIdx.append(a)
            elif self.agents[a].timescale == "steadystate":
                self.ssAgentIdx.append(a)
            elif self.agents[a].timescale == "dynamic":
                self.dynamicAgentIdx.append(a)
            else:
                raise ValueError(f'Observation type not supported for agent{i}!')

        self.marketAgentNum = len(self.marketAgentIdx)
        self.ssAgentNum = len(self.ssAgentIdx)
        self.dynamicAgentNum = len(self.dynamicAgentIdx)

        # initialize agents
        for a in self.agents:
            a.reset(self)

    # get measurements needed for an agent, specified in the fields of the agent class
    def take_sample(self, idx):
        all_sample = {}
        for i in self.agents[idx].obs:
            if i in ['Vseq', 'VLN', 'VLL']:
                ob = self.case.get_bus_V(self.agents[idx].bus1, i, self.agents[idx].phases)
            elif i in ['Iseq', 'Iph']:
                lineName = self.case.lineNames[self.agentsLineIndex[idx]]
                ob = self.case.get_line_I(lineName, i, self.agents[idx].phases)
            else:
                raise ValueError(f'Observation type not supported for agent{i}!')
                
            all_sample[i] = ob
            
        return all_sample

    # scale renewable generator and demands based on profile
    def apply_profile(self, rowIndex = None):
        
        # apply load changes
        loadC = self.loadProfile.at[rowIndex, 'value']
        loads = self.case.ckt.Loads
        loadNum = loads.Count
        
        # start from the first load
        loads.First
        for i in range(loadNum):
            loadP = loads.kW
            loadQ = loads.kvar
            loads.kW = loadP * loadC
            loads.kvar = loadQ * loadC
            loads.Next

        # DERs
        if self.DEREnable:
            # profiles percentage
            pvC = self.pvProfile.at[rowIndex, 'value']
            windC = self.windProfile.at[rowIndex, 'value']

            # set the generator dataframe
            self.case.genDF.loc[self.case.genDF["fuel"]=="pv", "Pmax"] = self.case.genDF.loc[self.case.genDF["fuel"]=="pv", "kVAbase"] * pvC
            self.case.genDF.loc[self.case.genDF["fuel"]=="wind", "Pmax"] = self.case.genDF.loc[self.case.genDF["fuel"]=="wind", "kVAbase"] * windC

            # put into model
            self.case.sync_gen_df()
            
        
    # reset the environment and start a market interval
    def new_timestamp(self, marketFlag = True, row_idx = None):
        # set demand DER max power (if there are DERs)
        self.apply_profile(row_idx)

        # if this timestep is a market clearance step
        if marketFlag:
            # use market to determine generator dispatch
            self.market.update_case(self.case)
            self.market.get_dispatch()

            # set generators in model
            self.case.sync_gen_df()

        # solve the model
        self.case.txt.Command = "set maxcontroliter=100"
        self.case.txt.Command = "set mode=snap"
        self.case.txt.Command = "Solve"

        assert self.case.sol.Converged, "Steady-state PF Failed!"

        # market agents can still observe even not taking actions
        for a in self.marketAgentIdx:
            self.agents[a].observe()

        # steady state agents
        for a in self.ssAgentIdx:
            self.agents[a].observe()
            self.agents[a].getAction()
            self.agents[a].setAction()

        # do dynamic simulation if there are dynamic agents
        if self.dynamicAgentNum > 0:
            self.reset_dynamic_event()            


    # initialize an episode of dynamic event simulation
    def reset_dynamic_event(self):
        # model is already initialized, get a random event
        self.fault = self.case.random_fault()
        self.case.txt.Command = self.fault.cmd

        self.currStep = 1

        # set dynamic mode
        self.case.txt.Command = "Solve mode=dynamics number=1 stepsize=" + str(self.ts)
        assert self.case.sol.Converged, "Dynamic PF Failed!"

        # get new observation for agents 
        for a in self.dynamicAgentIdx:
            self.agents[a].observe()
        

    # step through a dynamic event
    def step_dynamic_event(self):
        done = 0
        # check for max simulation time
        if self.currStep == self.maxStep:
            done = 1

        # actions of agents
        for a in self.dynamicAgentIdx:
            self.agents[a].getAction()
            self.agents[a].setAction()

        # solve this timestep
        self.currStep += 1
        self.case.solve_case()

        # rewards
        for a in self.dynamicAgentIdx:
            self.agents[a].getReward()

        # get new observation for agents 
        for a in self.dynamicAgentIdx:
            self.agents[a].observe()

        return done

    # step through the profile
    def step_ss_snapshots(self):
        done = 0
        # check for max simulation time
        if self.currStep == self.maxStep:
            done = 1

        # action of agents
        for a in self.ssAgentIdx:
            self.agents[a].getAction()
            self.agents[a].setAction()

        # solve time step
        self.currStep += 1
        self.case.solve_case()

        # rewards
        for a in self.ssAgentIdx:
            self.agents[a].getReward()

        for a in self.ssAgentIdx:
            self.agents[a].observe()

        return done

    # select a row from profiles
    def set_profile_row(self, mode):
        # shared hours in the profiles
        if self.DEREnable:
            common_index = self.loadProfile.index.intersection(self.pvProfile.index).intersection(self.windProfile.index)
        else:
            common_index = self.loadProfile.index
            
        # choose an hour
        if mode == 'sequential':
            if self.hourNum == None:
                self.hourNum = 0
            else:
                self.hourNum += 1
        elif mode == 'random':
            self.hourNum = np.random.choice(range(common_index.size))
        else:
            raise ValueError(f'Undefined profile indexing mode!')

        self.profileIndex = common_index[self.hourNum]
                

    ## Gym standard functions
    # reset a simulation and provide training agent state
    # mode: sequential/random
    def reset(self, mode='sequential'):

        # get a row number in profile
        self.set_profile_row(mode)

        # determine if need to run market
        if mode == 'sequential':
            if self.market.isMarketStep():
                marketFlag = 1
            else:
                marketFlag = 0

        # simulate the current profile
        self.new_timestamp(marketFlag, self.profileIndex)
        

        # return the observation of the agent under training
        if not self.trainingAgentIdx == None:
            ob_act = self.agents[self.trainingAgentIdx].state
        else:
            ob_act = None

        return ob_act



    # advancee step and return training agent state/reward
    def step(self):
        R = 0

        # if dynamic simulation, step in dynamic mode
        if len(self.dynamicAgentIdx) > 0:
            done = self.step_dynamic_event()
        else:
            done = self.step_ss_snapshots()

        ob_act = self.agents[self.trainingAgentIdx].state
        R = self.agents[self.trainingAgentIdx].reward
            

        return ob_act, R, done, {"Agent":self.trainingAgentIdx}
        
        
        
