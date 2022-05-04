import sys, os
sys.path.append(r"..\\")
import AI4Dist.env as aienv
from AI4Dist.market.market_template import DoNothingMarket
from AI4Dist.agent.OCRelayAgent import OCAgent
import matplotlib.pyplot as plt


case_path = r'C:\Users\Dongqi Wu\OneDrive\Work\iEnergy\case\IEEE34\ieee34Mod1.dss'
bus_coords = r'C:\Users\Dongqi Wu\OneDrive\Work\iEnergy\case\IEEE34\IEEE34_BusXY.csv'


params = {'time_step' : 0.0167,
          'max_step' : 600,
          'DEREnable' : True,
          'pv_profile' : r'..\..\profiles\pv_profile\ercot_houston_pv.csv',
          'wind_profile' : r'..\..\profiles\wind_profile\ercot_houston_wind.csv',
          'load_profile' : r'..\..\profiles\load_profile\ercot_houston_load.csv',
          }


# OC agents
agent1 = OCAgent('800','802','l1', 70, 0.1, 'IEEE-VIT')
agent2 = OCAgent('830','854','l15', 60, 0.1, 'IEEE-VIT')
agents = [agent1, agent2]

# no market for now
myMarket = DoNothingMarket()

# setup environment                
myEnv = aienv.env(case_path, agents, market=myMarket, params=params)

# do simulation
myEnv.reset('sequential')
done = 0
while not done:
    _, _, done, _ = myEnv.step()


plt.plot(agent1.waveform['I'])
plt.show()
