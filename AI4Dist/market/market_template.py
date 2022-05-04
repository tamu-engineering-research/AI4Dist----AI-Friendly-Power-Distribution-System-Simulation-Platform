

# basic template of a market
class market():
    def __init__(self):
        return


    def update_case(self, dssCase):
        self.case = dssCase


    def get_dispatch(self):
        return


    def isMarketStep(self):
        return 0



# do nothing market
class DoNothingMarket(market):
    pass

