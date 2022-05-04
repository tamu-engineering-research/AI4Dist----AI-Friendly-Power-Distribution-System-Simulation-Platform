import sys
sys.path.append(r"C:\Users\Dongqi Wu\OneDrive\Work\iEnergy")
import AI4Dist.env as aienv
from AI4Dist.market.market_template import DoNothingMarket
from AI4Dist.agents.OCRelayAgent import OCAgent
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import win32com.client


case_path = r'C:\Users\Dongqi Wu\OneDrive\Work\PRRL\HICSS\case\IEEE34\ieee34Mod1.dss'
bus_coords = r'C:\Users\Dongqi Wu\OneDrive\Work\PRRL\HICSS\case\IEEE34\IEEE34_BusXY.csv'


params = {'time_step' : 0.0167,
          'max_step' : 600,
          'DEREnable' : True,
          'pv_profile' : r'C:\Users\Dongqi Wu\OneDrive\Work\PRRL\HICSS\code\data\pv_profile\ercot_houston_pv.csv',
          'wind_profile' : r'C:\Users\Dongqi Wu\OneDrive\Work\PRRL\HICSS\code\data\wind_profile\ercot_houston_wind.csv',
          'load_profile' : r'C:\Users\Dongqi Wu\OneDrive\Work\PRRL\HICSS\code\data\load_profile\ercot_houston_load.csv',
          }



agents = []
myMarket = DoNothingMarket()
pd.set_option('display.max_columns', 1000)
myEnv = aienv.env(case_path, agents, market=myMarket, params=params)

