# define an abstract class template of the Agent
class agent():
    def __init__(self, bus1=None, bus2=None, timescale=None):
        # name (string) of the buses
        self.bus1 = bus1
        self.bus2 = bus2
        self.timescale = timescale
        self.env = None
        self.model = self.build()
        self.obs = None
        self.phases = 3
        self.svNum = None
        self.actNum = None
        self.trainable = False
        self.rewardFcn = None

    # reset internal states and assign the environment handle
    def reset(self, env):
        self.env = env
    
    # gather required observations from the environment
    def observe(self):
        pass

    # compute raw action using agent model
    def getAction(self):
        pass

    # modify the circuit model based on computed action
    def setAction(self):
        pass

    # (ML agent only) the environment should be set to train this agent, train the model if applicable
    def train(self):
        pass

    # (RL agent only) compute a reward corresponding to the action
    def getReward(self):
        pass
    
    # (ML agent only) construct a model, called at agent initialization
    def build(self):
        pass

    # (ML agent only) save model or weight to local drive
    def save(self):
        pass

    # (ML agent only) load saved model or weight stored previously
    def load(self):
        pass
