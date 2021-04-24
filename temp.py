import pandas as pd
class Fetcher:
    lastfetchdate=0
    listoftickerforthis=[] #List of stocks in this class for this db to be calulated the moment db is switched
    
    def getData(inst,days):
        pass
    def ABimport(AmiBroker,dfa,tempfile):
        dfa.to_csv(path, index=False,header=None)
        AmiBroker.Import(0, tempfile, "amicom.format")
        AmiBroker.RefreshAll()
        pass
    def importTicker():
        pass
    def RT():
        pass
    def quickRefresh():
        pass
    def backFill(inst,days):
       ABimport( pd.DataFrame(data= getData(inst,days) ).transpose() )


        pass
    def backFillAll():
        pass
    def backfillCurrent():
        pass
