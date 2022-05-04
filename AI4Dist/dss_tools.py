import win32com.client
import networkx as nx
import matplotlib.pyplot as plt
import csv
import numpy as np
import pandas as pd
import random
from .utils import *
from rl.core import Env


# load a DSS case into the dssCase object
# The dssCase object provide a easy interface with OpenDSS and
# utility functions
class dssCase():
    def __init__(self, case_path, ts):
        # initialize DSS interface objects
        self.dss_handle = win32com.client.Dispatch("OpenDSSEngine.DSS")
        self.txt = self.dss_handle.Text
        self.ckt = self.dss_handle.ActiveCircuit
        self.sol = self.ckt.Solution
        self.ActElmt = self.ckt.ActiveCktElement
        self.ActBus = self.ckt.ActiveBus
        
        # load the case passed through argument
        self.case_path = case_path
        self.load_case()
        self.solve_case()
        self.check_grounding()

        # examine the network info of the DSS case and create a graph
        self.get_network_info()
        self.create_graph()
        self.sort_edges()
        self.clrDict =[ 'sandybrown' for _ in range(self.graph.number_of_nodes()) ]
        self.edgeClrDict = [ 'grey' for _ in range(self.graph.number_of_edges()) ]
        self.sizeDict = np.ones(self.busNum) * 20

        # simulation information
        self.ts = ts

    # clean DSS memory    
    def reset_dss(self):
        self.dss_handle.ClearAll()

    # load a local case from the case file
    def load_case(self):
        self.txt.Command = f"compile [{self.case_path}]"

    # solve the current loaded case
    def solve_case(self):
        self.sol.Solve()

    # process case and get network informations
    def get_network_info(self):
        # list of bus names
        self.busNames = self.ckt.AllBusNames
        self.busNum = len(self.busNames)
        # list of phases for each bus
        self.busPhases = []
        for n in self.busNames:
            self.ckt.SetActiveBus(n)
            self.busPhases.append(self.ActBus.Nodes)

        # list of lines
        self.lineNames = self.ckt.Lines.AllNames
        self.lineNum = self.ckt.Lines.Count

        self.lineT = []
        for n in self.lineNames:
            full_name = 'line.' + n
            self.ckt.SetActiveElement(full_name)
            F = self.ActElmt.Properties('Bus1').val.split('.')[0]
            T = self.ActElmt.Properties('Bus2').val.split('.')[0]
            
            # take only the 3-phase bus name
            try:
                self.lineT.append((self.busNames.index(F),self.busNames.index(T)))
            except Exception:
                print("Inconsistency in Bus/Line declearation!")
            
        # add transformers as lines (for graph making purpose)
        self.xfmrName = self.ckt.Transformers.AllNames
        self.xfmrNum = self.ckt.Transformers.Count
        self.xfmrT = []
        for tr in self.xfmrName:
            full_name = 'Transformer.' + tr
            self.ckt.SetActiveElement(full_name)
            F = self.busNames.index(self.ActElmt.busNames[0].split('.')[0])
            T = self.busNames.index(self.ActElmt.busNames[1].split('.')[0])

            self.xfmrT.append((F,T))
        self.xfmrT = list(set(self.xfmrT))

        # check generators
        self.parse_generators()

    # check and store existing generators in the system
    def parse_generators(self):
        # containers
        genFields = ["name", "fuel", "bus", "Pg", "Qg", "Vg", "kVAbase", "Pmin", "Pmax", "Qmin", "Qmax", "Pcost"]
        genDict = {key: [] for key in genFields}

        # source as inf bus, DSS only allow one Vsource now?
        self.ckt.Vsources.First
        self.ckt.SetActiveElement(f'vsource.{self.ckt.Vsources.Name}')
        genDict['name'].append(self.ckt.Vsources.Name)
        genDict['Vg'].append(self.ckt.Vsources.pu)
        genDict['fuel'].append('source')
        genDict['bus'].append(self.ckt.ActiveCktElement.BusNames[1])
        genDict['kVAbase'].append(1e9)
        genDict['Pmin'].append(-1e9)
        genDict['Pmax'].append(1e9)
        genDict['Qmin'].append(-1e9)
        genDict['Qmax'].append(1e9)
        genDict['Pg'].append(0)
        genDict['Qg'].append(0)
        genDict['Pcost'].append(0)
        
        # rotational
        gens = self.ckt.Generators
        self.genNum = gens.Count
        gens.First
        for i in range(self.genNum):
            self.ckt.SetActiveElement(f'generator.{gens.Name}')
            genDict['name'].append(gens.Name)
            genDict['Vg'].append(gens.kV)
            genDict['fuel'].append('default')
            genDict['bus'].append(self.ckt.ActiveCktElement.BusNames[0])
            genDict['kVAbase'].append(gens.kVArated)
            genDict['Pmin'].append(-1e9)
            genDict['Pmax'].append(1e9)
            genDict['Qmin'].append(-1e9)
            genDict['Qmax'].append(1e9)
            genDict['Pg'].append(gens.kW)
            genDict['Qg'].append(gens.kvar)
            genDict['Pcost'].append(0)
            gens.Next

        # pv generators
        pvs = self.ckt.PVSystems
        self.pvNum = pvs.Count
        pvs.First
        for i in range(self.pvNum):
            self.ckt.SetActiveElement(f'pvsystem.{pvs.Name}')
            genDict['name'].append(pvs.Name)
            genDict['Vg'].append(self.ckt.ActiveCktElement.SeqVoltages[1])
            genDict['fuel'].append('pv')
            genDict['bus'].append(self.ckt.ActiveCktElement.BusNames[0])
            genDict['kVAbase'].append(pvs.kVArated)
            genDict['Pmin'].append(-1e9)
            genDict['Pmax'].append(1e9)
            genDict['Qmin'].append(-1e9)
            genDict['Qmax'].append(1e9)
            genDict['Pg'].append(pvs.kW)
            genDict['Qg'].append(pvs.kvar)
            genDict['Pcost'].append(0)
            pvs.Next
        
        self.genDF = pd.DataFrame.from_dict(genDict)

    # add generator into the system
    def add_gen(self, **kwargs):
        genFields = ["name", "fuel", "bus", "Pg", "Qg", "Vg", "kVAbase", "Pmin", "Pmax", "Qmin", "Qmax", "Pcost"]
        newGen = {}
        newGen = {key: 0 for key in genFields}
        # add to gen dataframe
        for k, v in kwargs.items():
            newGen[k] = v
            
        self.genDF = self.genDF.append(newGen, ignore_index=True)
        # add to network model
        if newGen['fuel'] == 'pv':
            genS = np.linalg.norm([newGen["Pg"], newGen["Qg"]])
            genPF = newGen["Pg"] / genS
            self.txt.Command = f'New pvsystem.{newGen["name"]} phase=1 bus1={newGen["bus"]} kv={newGen["Vg"]} kva={genS} pf={genPF} pctPmpp=100'
        else:
            self.txt.Command = f'New generator.{newGen["name"]} bus1={newGen["bus"]} kv={newGen["Vg"]} kw={newGen["Pg"]} kva={newGen["kVAbase"]} model=7 Balanced=Yes'
        

    # sync data a generator DF with model
    def sync_gen_df(self):
        # 1st (source) is always slack
        for i in range(1, self.genDF.shape[0]):
            gen = self.genDF.iloc[i]
            if gen["fuel"] == 'pv':
                genS = np.linalg.norm([gen["Pg"], gen["Qg"]])
                genPF = gen["Pg"] / genS
                self.txt.Command = f'edit pvsystem.{gen["name"]} kv={gen["Vg"]} kva={genS} pf={genPF}'
            else:
                self.txt.Command = f'edit generator.{gen["name"]} kv={gen["Vg"]} kw={gen["Pg"]} kvar={gen["Qg"]}'
        
        return
    
    # check if ground current path exist across the case
    def check_grounding(self):
        self.groundPath = False
        # go through transformers, loads and capactiors
        xfmrs = self.ckt.Transformers
        xfmrs.First
        for i in range(xfmrs.Count):
            if not xfmrs.IsDelta:
                self.groundPath = True
                return
            xfmrs.Next
            
        # capacitors
        caps = self.ckt.Capacitors
        caps.First
        for i in range(caps.Count):
            if not caps.IsDelta:
                self.groundPath = True
                return
            caps.Next
            
        # loads
        loads = self.ckt.Loads
        loads.First
        for i in range(loads.Count):
            if not loads.IsDelta:
                self.groundPath = True
                return
            loads.Next       
        
        

    ## GRAPH-RELATED FUNCTIONS
    # create a graph for the network using networkx
    def create_graph(self):
        # create new un-directed graph first
        self.graph = nx.Graph()
        
        # add lines as edges of this graph
        for l in self.lineT:
            self.graph.add_edge(l[0], l[1])

        # add transformers as edges of this graph
        for t in self.xfmrT:
            self.graph.add_edge(t[0], t[1])

        # change to directed graph and remove edges according to radial structure
        self.graph = self.graph.to_directed()

        ## for every line between bus, remove the backward edge that is not the assumed positive direction
        # compute the distance from source
        dist_from_source = nx.single_source_shortest_path_length(self.graph, 0) 

        # lines
        for l in self.lineT:
            # if bus1 is closer to souce than bus2
            if dist_from_source[l[0]] < dist_from_source[l[1]]:
                # remove the edge (bus2 -> bus1)
                try:
                    self.graph.remove_edge(l[1], l[0])
                except:
                    print(f'Warning: line {self.busNames[l[1]]} -> {self.busNames[l[0]]} cannot be removed, please check graph!')

        # transformers
        for t in self.xfmrT:
            # if bus1 is closer to souce than bus2
            if dist_from_source[t[0]] < dist_from_source[t[1]]:
                # remove the edge (bus2 -> bus1)
                try:
                    self.graph.remove_edge(t[1], t[0])
                except:
                    print(f'Warning: line {self.busNames[t[1]]} -> {self.busNames[t[0]]} cannot be removed, please check graph!')
        


    # draw the network graph using matplotlib
    def draw_graph(self):
        plt.figure()
        self.mark_fuse_lines()
        nx.draw(self.graph, pos=self.posDict, node_color=self.clrDict,\
                edge_color=self.edgeClrDict, with_labels=False, node_size=self.sizeDict)
        
        plt.show()
        

    # sort nodes using DFS
    def sort_edges(self):
        self.edge_order = list(nx.dfs_edges(self.graph, source=0))


    # read bus coordinates from external file
    def read_bus_coords(self, fp):
        fh = open(fp, "r")
        allLines = fh.readlines()
        # guess the delimiter as there is no common format
        sn = csv.Sniffer()
        delim = sn.sniff(allLines[0]).delimiter
        
        # each line contains [busName, x, y] separated by space (?)
        self.posDict = {}
        for line in allLines:
            name, x, y = line.split(delim)
            busID = self.busNames.index(name)
            self.posDict.update({busID:[getNum(x),getNum(y)]})
            
    # mark the nodes that have protection problem as red
    def mark_logged_buses(self, log):

        # all bus where a fault is not detected
        allBuses = [self.busNames.index(l.fault.bus) for l in log]

        # paint these buses in red
        for n in self.graph.nodes:
            if n in allBuses:
                self.clrDict[n] = 'red'
                self.sizeDict[n] = 75
            else:
                self.clrDict[n] = 'sandybrown'
                self.sizeDict[n] = 10

    # mark the lines with fuse as yellow
    def mark_fuse_lines(self):

        # find the list of all fuses
        fuseLines = []
        fuses = self.ckt.Fuses
        fuseNum = fuses.Count
        fuses.First
        for i in range(fuseNum):
            fuseLines.append(fuses.MonitoredObj.replace('line.',''))
            fuses.Next

        # find the 2nd bus of the line
        for line in fuseLines:
            lineidx = self.lineNames.index(line)
            bus2 = self.lineT[lineidx][1]
            self.clrDict[bus2] = 'blue'
            self.sizeDict[bus2] = 50
    
        
    # get line current measurement using line name
    def get_line_I(self, name, field, phase):
        full_name = 'line.' + name
        self.ckt.SetActiveElement(full_name)
        if phase == 3:
            if field == 'Iseq':
                res = self.ActElmt.SeqCurrents[0:3]
            elif field == 'Ipuseq':
                res = self.ActElmt.SeqCurrents[0:3]
            elif field == 'Iph':
                res = cart_to_pol(self.ActElmt.Currents[0:6])
            else:
                raise ValueError(f'Please use a valid field name for Line measurement')
        # return both phase currents if 2 conductors
        elif phase == 2:
            res = cart_to_pol(self.ActElmt.Currents[0:4])
            
        # return single phase current if only 1 conductor
        elif phase == 1:
            res = cart_to_pol(self.ActElmt.Currents[0:2])
            
        return res
    
        
    # get bus voltage measurement using busname
    # returns:
    # [mag, angle]
    # mag = [magA, magB, magC]
    # angle = [angleA, angleB, angleC]
    # or
    # Sequence
    # [V0, V1, V2]
    def get_bus_V(self, name, field, phase):
        self.ckt.SetActiveBus(name)
        baseV = self.ActBus.kVbase * 1000
        if phase == 3:       
            if field == 'Vseq':
                res = self.ActBus.SeqVoltages
            if field == 'Vpuseq':
                res = [i / baseV for i in self.ActBus.SeqVoltages]
            elif field == 'VLN':
                mag, angle = cart_to_pol(self.ActBus.Voltages)
                res = [mag, angle]
            elif field == 'VLL':
                mag, angle = cart_to_pol(self.ActBus.VLL)
                res = [mag, angle]
            else:
                raise ValueError(f'Please use a valid field name for Bus measurement')

        elif phase == 2:
            if field == 'VLN':
                mag, angle = cart_to_pol(self.ActBus.Voltages)
                res = [mag, angle]
            elif field == 'VLL':
                mag, angle = cart_to_pol(self.ActBus.VLL)
                res = [mag, angle]
            else:
                raise ValueError(f'Please use a valid field name for Line measurement') 
    
        elif phase == 1:
            if field == 'VLN':
                mag, angle = cart_to_pol(self.ActBus.Voltages)
                res = [mag, angle]
            elif field == 'VLL':
                mag, angle = cart_to_pol(self.ActBus.VLL)
                res = [mag, angle]
            else:
                raise ValueError(f'Please use a valid field name for Line measurement')   

        
        return res

    # edit a property, or a list of properties, of a DSS element
    def edit_elmt(self, name, fields, vals):
        cmd = f'Edit {name}'
        # if providing a list of properties, iterate through
        if isinstance(fields, list):
            for f, v in zip(fields, vals):
                cmd += f' {f}={v}'
        # if only one property, add and execute
        else:
            cmd += f'{fields}={vals}'

        self.txt.Command = cmd

    # trip an element in the netwrok
    def trip_elmt(self, elmt):
        self.txt.Command = f'open line.{elmt} term=1'

    
    # create a random fault in this case
    def random_fault(self):
        randFault = fault(self.busNames, self.busPhases, self.ts, self.groundPath)
        
        return randFault

    # create a random transient event in this case
    def random_event(self):
        pass
        
    

# fault class for DSS
class fault():
    def __init__(self, buses, phases, ts, GNDFlag):
        self.GNDFlag = GNDFlag
        self.bus = self.rand_bus(buses[2:], phases[2:])
        self.phases = self.rand_phase(buses[2:], phases[2:])
        self.R = self.rand_resistance()
        self.T = self.rand_time(ts)
        self.cmd = self.get_cmd_string()

    # location of fault        
    def rand_bus(self, buses, phases):
        # return a random bus in the system
        self.bus_idx = np.random.choice(range(len(buses)))
        if self.GNDFlag:
            # 1 or 3-phase buses if GND path exists
            while not (len(phases[self.bus_idx])==1 or len(phases[self.bus_idx])==3):
                self.bus_idx = np.random.choice(range(len(buses)))
        else:
            # only 3-phase buses if GND path does not exist
            while not len(phases[self.bus_idx])==3:
                self.bus_idx = np.random.choice(range(len(buses)))            
        
        return buses[self.bus_idx]

    # return a fault type
    def rand_phase(self, buses, phases):
        p = phases[self.bus_idx]

        # if 1p line, only SLG  possible 
        if len(p) == 1:
            self.type = '1'
            return str(p[0])

        # if 2p line, SLG, LL or LLG
        if len(p) == 2:
            if self.GNDFlag:
                self.type = np.random.choice(['1','2'])
            else:
                self.type = '2'
                
            if self.type == '1':
                return np.random.choice(p)
            elif self.type == '2' or self.type == '2g':
                return np.random.choice(p, 2, replace=False)
        
        # if 3p line, can have all kinds of fault
        elif len(p) == 3:
            if self.GNDFlag:
                self.type = np.random.choice(['1','2','3'])
            else:
                self.type = np.random.choice(['2','3'])
                
            if self.type == '1':
                return np.random.choice(['1','2','3'])
            elif self.type == '2' or self.type == '2g':
                return np.random.choice(['1','2','3'], 2, replace=False)
            else:
                return ['1','2','3']

        
    def rand_resistance(self):
        # corresponding to low, med, high res fault
        fault_r_range = [[0.002,0.01],[0.01, 0.1],[0.1,1],[1,30]]
        fault_r = fault_r_range[np.random.choice([0,1,2,3])]
        #fault_r = fault_r_range[0]
        R = np.random.uniform(fault_r[0],fault_r[1])
        R = round(R, 4)
        return R

    def rand_time(self, ts):
        return round(((np.floor(np.random.uniform(15, 30))+0.1) * ts), 4)
        #return round((4.1 * ts), 4)

    # generate DSS command string from randomized attributes
    def get_cmd_string(self):
        cmd = 'New Fault.F1 '
        # number of phases
        cmd += 'Phases=' + str(len(self.phases))
        # format the faulted lines to the input form
        if self.type == '1':
            cmd += ' Bus1=' + self.bus + '.' + self.phases[0]
        elif self.type == '2':
            cmd += ' Bus1=' + self.bus + '.' + self.phases[0] + '.0'
            cmd += ' Bus2=' + self.bus + '.' + self.phases[1] + '.0'
        elif self.type == '2g':
            cmd += ' Bus1=' + self.bus + '.' + self.phases[0] + '.' + self.phases[0]
            cmd += ' Bus2=' + self.bus + '.' + self.phases[1] + '.0'
        elif self.type == '3':
            cmd += ' Bus1=' + self.bus + '.1.2.3'
        # fault resistance
        cmd += ' R=' + str(self.R)
        # fault time
        cmd += ' ONtime=' + str(self.T)

        return cmd
