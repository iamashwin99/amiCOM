import pandas as pd
import os
class exchange():
    
    # A class with methods to import data and to feed data from amibroker
    # Has the methods:
    #     ImportTickers : Import tickers from the exchange and store it to tickerlist file
    #     ImportCurrent : Import the data for the current symbol in amibroker
    #     LiveRefersh: Import current market price for all symbols
    #     BackfillAll: Import data for all the tickers in the db from a give time range
    #     ConvertTicker2ExhcangeID : Convert the amibroker ticker (as in the list and in the progarm) to the exchange id (used to fetch)
    #     ConvertAmiID2ExchangeID : Convert the exchange id to the amibroker ticker
    #     Pushdf2Amibroker : push the ohlc df to a file then import it to amibroker
    #     onCloseDB: things to be done before closing/ switching db
    # Has the following data :
    #     db : a pandas document of OHLC values for the given import session
    #     dbaselocation: Location of the Amibroker DB
    #     dblistlocation: location of the tickers list 
    lastfetchdate=0
    listoftickerforthis=[] #List of stocks in this class for this db to be calulated the moment db is switched
    df=pd.DataFrame()
    tempFilePath = ""
    def __init__(self,dbname,dbaselocation,dblistlocation):
        # if dbaselocation dosent exists create it
        if not os.path.exists(dbaselocation):
            os.makedirs(dbaselocation)
        #if file dblistlocation dosent exists create it
        if not os.path.exists(dblistlocation):
            #create an emty file in dblistlocation
            open(dblistlocation,'w').close()
        pass
    def Pushdf2Amibroker(self,AmiBroker):
        df.to_csv(tempFilePath, index=False,header=None)
        AmiBroker.Import(0, tempfile, "amicom.format")
        AmiBroker.RefreshAll()
    def ImportTickers(self):
        pass
    def ImportCurrent(self):
        pass
    def LiveRefersh(self):
        pass
    def BackfillAll(self):
        pass
    def ConvertTicker2ExhcangeID(self,inst):
        pass
    def ConvertAmiID2ExchangeID(self,inst):
        pass
    def onCloseDB(self):
        pass
    


    

    
    