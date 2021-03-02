import datetime  # For datetime objects
import os  # To manage paths
import sys
from typing import NoReturn
from win32com.client import Dispatch
import datetime 
from sys import argv
import tkinter
import time
import pandas as pd
import yfinance as yf
import logging
import re
import tkinter.messagebox as tkMessageBox
# ttk makes the window look like running Operating System’s theme
from tkinter import ttk
import tkinter.scrolledtext as st 
import random
from jugaad_data.nse import NSELive
n = NSELive()
from keys import *
lastClose = 0
abDatabase = 'C:\\amiCOM\\DB'
NIFTY50DB = 'C:\\amiCOM\\DB\\NIFTY50'
NIFTY100DB = 'C:\\amiCOM\\DB\\NIFTY100'
NIFTY200DB = 'C:\\amiCOM\\DB\\NIFTY200'
CUSTOM1DB = 'C:\\amiCOM\\DB\\CUSTOM1'
NEAREXPDB = 'C:\\amiCOM\\DB\\NEAREXP'
BNBDB = 'C:\\amiCOM\\DB\\BINANCE'

TempFile= 'C:\\amiCOM\\temp.txt'
open(TempFile, 'w').close() # Clear temp file while first load

NIFTY50list = 'C:\\amiCOM\\TickerList\\NIFTY50.txt'
NIFTY100List = 'C:\\amiCOM\\TickerList\\NIFTY100.txt'
NIFTY200List = 'C:\\amiCOM\\TickerList\\NIFTY200.txt'
CUSTOM1List = 'C:\\amiCOM\\TickerList\\CUSTOM1.txt'
NEAREXPList = 'C:\\amiCOM\\TickerList\\NEAREXP.txt'
BNBList = 'C:\\amiCOM\\TickerList\\BINANCE.txt'

LOGDIR = 'C:\\amiCOM\\Logs.txt'

indicesY = ['^NSEI',
 '^NSMIDCP',
 '^CNX100',
 '^CNX200',
 '^CNX500',
 '^NSEMDCP50',
 '^CRSMID',
 '^CNXSC',
 '^INDIAVIX',
 'NETFMID150.NS',
 'NIFTYSMLCAP50.NS',
 'NIFTYSMLCAP250.NS',
 'MSL400.BO',
 '^NSEBANK',
 '^CNXAUTO',
 '^CNXFIN',
 '^CNXFIN',
 '^CNXFMCG',
 '^CNXIT',
 '^CNXMEDIA',
 '^CNXMETAL',
 '^CNXPHARMA',
 '^CNXPSUBANK',
 'NIFTYPVTBANK.NS',
 '^CNXREALTY',
 '^CNXDIVOP',
 'NI15.NS',
 'NIFTYQUALITY30.NS',
 'NV20.NS',
 'NIFTYTR2XLEV.NS',
 'NIFTYPR2XLEV.NS',
 'NIFTYTR1XINV.NS',
  'NIFTYPR1XINV.NS',
 '^NSEDIV',
 'na',
 'NFTY',
 'NIFTY100_EQL_WGT.NS',
 'NIFTY100LOWVOL30.NS',
 'NIFTY200QUALTY30.NS',
 'NIFTYALPHALOWVOL.NS',
 'NIFTY200MOMENTM30.NS',
 '^CNXCMDT',
 '^CNXCONSUM ',
 'CPSE.NS',
 '^CNXENERGY',
 '^CNXINFRA',
 'LIX15.NS',
 ' NIFTYMIDLIQ15.NS',
 '^CNXMNC',
 '^CNXPSE',
 '^CNXSERVICE',
 '^CNX100',
 'NIFTYGS8TO13YR.NS',
 'NIFTYGS10YR.NS',
 'NIFTYGS10YRCLN.NS ',
 'NIFTYGS4TO8YR.NS',
 'NIFTYGS11TO15YR.NS',
 'NIFTYGS11TO15YR.NS',
 'NIFTYGSCOMPOSITE.NS']
indicesN=['NIFTY 50',
 'NIFTY NEXT 50',
 'NIFTY 100',
 'NIFTY 200',
 'NIFTY 500',
 'NIFTY MIDCAP 50',
 'NIFTY MIDCAP 100',
 'NIFTY SMLCAP 100',
 'INDIA VIX',
 'NIFTY MIDCAP 150',
 'NIFTY SMLCAP 50',
 'NIFTY SMLCAP 250',
 'NIFTY MIDSML 400',
 'NIFTY BANK',
 'NIFTY AUTO',
 'NIFTY FIN SERVICE',
 'NIFTY FINSRV25 50',
 'NIFTY FMCG',
 'NIFTY IT',
 'NIFTY MEDIA',
 'NIFTY METAL',
 'NIFTY PHARMA',
 'NIFTY PSU BANK',
 'NIFTY PVT BANK',
 'NIFTY REALTY',
 'NIFTY DIV OPPS 50',
 'NIFTY GROWSECT 15',
 'NIFTY100 QUALTY30',
 'NIFTY50 VALUE 20',
 'NIFTY50 TR 2X LEV',
 'NIFTY50 PR 2X LEV',
 'NIFTY50 TR 1X INV',
 'NIFTY50 PR 1X INV',
 'NIFTY50 DIV POINT',
 'NIFTY ALPHA 50',
 'NIFTY50 EQL WGT',
 'NIFTY100 EQL WGT',
 'NIFTY100 LOWVOL30',
 'NIFTY200 QUALTY30',
 'NIFTY ALPHALOWVOL',
 'NIFTY200MOMENTM30',
 'NIFTY COMMODITIES',
 'NIFTY CONSUMPTION',
 'NIFTY CPSE',
 'NIFTY ENERGY',
 'NIFTY INFRA',
 'NIFTY100 LIQ 15',
 'NIFTY MID LIQ 15',
 'NIFTY MNC',
 'NIFTY PSE',
 'NIFTY SERV SECTOR',
 'NIFTY100ESGSECLDR',
 'NIFTY GS 8 13YR',
 'NIFTY GS 10YR',
 'NIFTY GS 10YR CLN',
 'NIFTY GS 4 8YR',
 'NIFTY GS 11 15YR',
 'NIFTY GS 15YRPLUS',
 'NIFTY GS COMPSITE']

open(LOGDIR, 'w').close() # Clear log file while first load
logging.basicConfig(format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')
#logging.basicConfig(filename=LOGDIR, level=logging.debug, format='%(asctime)s - %(levelname)s - %(message)s')
AmiBroker = Dispatch("Broker.Application")
AmiBroker.visible=True

if(Bapi_key !='BLABLA'):
    from binance.client import Client
    Bclient = Client(Bapi_key, Bapi_secret)


AmiBroker.LoadDatabase(BNBDB)




## Methods
def YahooOrNSE(inst):
    return bool(re.match(r"(^\^\w+|\w+.NS)",inst)) # return if ticker is of yahoo or not
def opti2inst(inst):
    return inst.split("-")[1]

def Convert2(dest,inst):
    if dest == 'y': #destination is yahoo
        if (not YahooOrNSE(inst)): # source is not already yahoo
            if inst not in indicesN: #check if inst is not indices
                return inst+'.NS'
            else:
                return indicesY[indicesN.index(inst)] #if indices replace it correctly
        else:
            return inst
    elif dest == 'n':
        if (YahooOrNSE(inst)): #ensure source is yahoo
            if inst not in indicesY:
                return inst.split('.')[0] # remove .NS
            else:
                return indicesN[indicesY.index(inst)] #if indices replace it correctly
        else:
            return inst.split('.')[0] #remove .bnb from bnb

def IsOption(inst):
    if(YahooOrNSE(inst)): # If ints yahoo then not option
        return 0
    else:
        return bool(re.match(r"^OPTI-",inst))
    
## Data filling methods
def ImportTickers():
    
    source=DB.get() #"NIFTY50" ,"NIFTY100", "NIFTY200", "CUSTOM1"
    ticker =[]
    if(source =="NIFTY50" ): 
        filename = NIFTY50list

    elif(source =="NIFTY100" ): 
        filename = NIFTY100List

    elif(source =="NIFTY200" ): 
        filename = NIFTY200List
    elif(source=="NEAREXP"): 
        filename = NEAREXPList
    elif(source=="BNBDB"): 
        import re
        filename = BNBList
        prices = client.get_all_tickers()
        simlist=[]
        for i in range (0,len(prices)):
            a=(prices)[i]["symbol"]
            if ( re.search('(\w+USDT)',a)) :
                simlist.append(a)
        with open(filename, 'w') as f: 
            for item in simlist:
                f.write(item+'.BNB\n')
                
        
    else:
        filename = CUSTOM1List

    with open(filename) as f:
        ticker = f.readlines()
        ticker = [x.strip() for x in ticker] # remove whitespace characters like `\n` at the end of each line
        ticker = [Convert2('n',x) for x in ticker]

    for count in range(0, len(ticker)):
        if not IsOption(ticker[count]):
            AmiBroker.Stocks.Add(Convert2('n',ticker[count])) # Add tickers from list
        else:
            setOptions(opti2inst(ticker[count]))
            
    Qty = AmiBroker.Stocks.Count
    for i in range(0, Qty):
            inst = AmiBroker.Stocks(i).Ticker
            if inst not in ticker and (not IsOption(inst)):
                AmiBroker.Stocks.Remove(inst) # remove tickers not in list

    AmiBroker.RefreshAll()
    AmiBroker.SaveDatabase()
    



def Backfill():
    #return 0


    days2Fill = int(daystofill.get()) 
    if days2Fill < 7:
        interval_length = '1m'
    elif days2Fill < 60:
        interval_length = '5m'
    else:
        interval_length = '1d'

    s_date = datetime.datetime.now()-datetime.timedelta(days = int(days2Fill))
    e_date =  datetime.datetime.now()+datetime.timedelta(days = 1)

    start_date = s_date.strftime("%Y-%m-%d")
    end_date = e_date.strftime("%Y-%m-%d")

    continous = 0
    while continous == 0:
        Qty = AmiBroker.Stocks.Count
        for i in range(0, Qty):
            inst = AmiBroker.Stocks(i).Ticker
            #logging.debug("Getting data for "+str(inst))
            tickerData = yf.Ticker(Convert2('y',inst))
            tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
            #logging.debug("Got data for "+str(inst))
            timelist = list(tickerDf.index)
            for count in range(0, len(tickerDf)):                
                asking_time = timelist[count].strftime('%d/%m/%Y %H:%M:%S')
                asking_open = tickerDf['Open'][count]
                asking_low = tickerDf['Low'][count]
                asking_high = tickerDf['High'][count]
                asking_close = tickerDf['Close'][count]
                asking_volume = tickerDf['Volume'][count]                
                ticker = AmiBroker.Stocks.Add(Convert2('n',inst))
                quote = ticker.Quotations.Add(asking_time)
                quote.Open = asking_open
                quote.Low = asking_low
                quote.High = asking_high
                quote.Close = asking_close
                AmiBroker.RefreshAll()

def ImportThreaded():
    if(DB.get()=="BNBDB"):
        BNBBackfill()
        return 0 
    #return 0
    #global daysToFil
    path = TempFile
    open(path, 'w').close()
    #file = open(path, 'w')
    Qty = AmiBroker.Stocks.Count
    days2Fill = int(daystofill.get()) 
    if days2Fill < 7:
        interval_length = '1m'
    elif days2Fill < 60:
        interval_length = '5m'
    else:
        interval_length = '1d'

    s_date = datetime.datetime.now()-datetime.timedelta(days = int(days2Fill))
    e_date =  datetime.datetime.now()+datetime.timedelta(days = 1)

    start_date = s_date.strftime("%Y-%m-%d")
    end_date = e_date.strftime("%Y-%m-%d")
    
    listofstocks=[]
    for i in range(0, Qty):
        ABstock = AmiBroker.Stocks(i).Ticker
        if(not IsOption(ABstock)):
            listofstocks.append(Convert2('y',ABstock))
        else:
            setOptions(opti2inst(ABstock))
            
    data = yf.download(" ".join(listofstocks), interval=interval_length, start=start_date, end=end_date,group_by = 'ticker',auto_adjust = True,threads = True)
    
    availableList=list(dict(data.keys()).keys()) #Looking for a better way !
    
    for i in range(0, len(availableList)):
        inst =  availableList[i]
        #logging.debug("Getting data for "+str(inst))
        #tickerData = yf.Ticker(inst)
        #tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
        tickerDf = data[inst]
        #logging.debug("Got data for "+str(inst))
        timelist = list(tickerDf.index)
        ticker=[inst]*len(tickerDf)
        ticker =[Convert2('n',x) for x in ticker]
        ymd = [ x.strftime('%Y%m%d') for x in timelist  ]
        time =  [ x.strftime('%H:%M') for x in timelist  ]
        asking_open = tickerDf['Open']
        asking_low = tickerDf['Low']
        asking_high = tickerDf['High']
        asking_close = tickerDf['Close']
        asking_volume = tickerDf['Volume']
        d = [ticker,ymd,time,asking_open,asking_high,asking_low,asking_close,asking_volume ]
        dfa = pd.DataFrame(data=d).transpose()
        dfa.to_csv(path, mode='a', index=False,header=None)
    
    AmiBroker.Import(0, path, "amicom.format")
    AmiBroker.RefreshAll()

def QuickImportThreaded():
     
    #return 0
    #global daysToFil
    path = TempFile
    open(path, 'w').close()
    #file = open(path, 'w')
    Qty = AmiBroker.Stocks.Count
    days2Fill = int(daystofill.get()) 
    if days2Fill < 7:
        interval_length = '1m'
    elif days2Fill < 60:
        interval_length = '5m'
    else:
        interval_length = '1d'

    s_date = datetime.datetime.now()-datetime.timedelta(days = 1)
    e_date =  datetime.datetime.now()+datetime.timedelta(days = 1)

    start_date = s_date.strftime("%Y-%m-%d")
    end_date = e_date.strftime("%Y-%m-%d")
    
    listofstocks=[]
    for i in range(0, Qty):
        ABstock = AmiBroker.Stocks(i).Ticker
        if(not IsOption(ABstock)):
            listofstocks.append(Convert2('y',ABstock))
        else:
            #setOptions(opti2inst(ABstock))
            pass
            

    data = yf.download(" ".join(listofstocks), interval=interval_length, start=start_date, end=end_date,group_by = 'ticker',auto_adjust = True,threads = True)
    
    availableList=list(dict(data.keys()).keys()) #Looking for a better way !
    
    for i in range(0, len(availableList)):
        inst = availableList[i]
        #logging.debug("Getting data for "+str(inst))
        #tickerData = yf.Ticker(inst)
        #tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
        tickerDf = data[inst]
        #logging.debug("Got data for "+str(inst))
        timelist = list(tickerDf.index)
        ticker=[inst]*len(tickerDf)
        ticker =[Convert2('n',x) for x in ticker]
        ymd = [ x.strftime('%Y%m%d') for x in timelist  ]
        time =  [ x.strftime('%H:%M') for x in timelist  ]
        asking_open = tickerDf['Open']
        asking_low = tickerDf['Low']
        asking_high = tickerDf['High']
        asking_close = tickerDf['Close']
        asking_volume = tickerDf['Volume']
        d = [ticker,ymd,time,asking_open,asking_high,asking_low,asking_close,asking_volume ]
        dfa = pd.DataFrame(data=d).transpose()
        dfa.to_csv(path, mode='a', index=False,header=None)
    
    AmiBroker.Import(0, path, "amicom.format")
    AmiBroker.RefreshAll()
    


def ImportCur():
    inst = AmiBroker.ActiveDocument.Name
    if(IsOption(inst)):
            setOptions(opti2inst(inst))
            return 0

    open(TempFile, 'w').close()

    days2Fill = int(daystofill.get()) 
    if days2Fill < 7:
        interval_length = '1m'
    elif days2Fill < 60:
        interval_length = '5m'
    else:
        interval_length = '1d'

    s_date = datetime.datetime.now()-datetime.timedelta(days = int(days2Fill))
    e_date =  datetime.datetime.now()+datetime.timedelta(days = 1)

    start_date = s_date.strftime("%Y-%m-%d")
    end_date = e_date.strftime("%Y-%m-%d")

     
    logMe("Getting data for "+str(inst))
    tickerData = yf.Ticker(Convert2('y',inst))
    tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
    logMe("Got data for "+str(inst))
    timelist = list(tickerDf.index)
    ticker=[inst]*len(tickerDf)
    ticker =[Convert2('n',x) for x in ticker]
    ymd = [ x.strftime('%Y%m%d') for x in timelist  ]
    time =  [ x.strftime('%H:%M') for x in timelist  ]
    asking_open = tickerDf['Open']
    asking_low = tickerDf['Low']
    asking_high = tickerDf['High']
    asking_close = tickerDf['Close']
    asking_volume = tickerDf['Volume']
    d = [ticker,ymd,time,asking_open,asking_high,asking_low,asking_close,asking_volume ]
    dfa = pd.DataFrame(data=d).transpose()
    dfa.to_csv(TempFile, index=False,header=None)
    AmiBroker.Import(0, TempFile, "amicom.format")
    AmiBroker.RefreshAll()

def refreshOPtions():
    setOptions("NIFTY")
    setOptions("BANKNIFTY")

def setOptions(inst):
    option_chain = n.index_option_chain(inst)
    if len(option_chain)==0:
        return 0
    now=datetime.datetime.now() 
    list=[]
    for i in range(0,len(option_chain['filtered']['data'])):
        side ='CE'
        name='OPTI-'+str(inst)+'-'+str(option_chain['filtered']['data'][i][side]['strikePrice'])+side
        date =now.strftime('%Y%m%d')
        time = now.strftime('%H:%M')
        price = option_chain['filtered']['data'][i][side]["lastPrice"]
        volume=abs(option_chain['filtered']['data'][i][side]["changeinOpenInterest"])
        openint=abs(option_chain['filtered']['data'][i][side]["openInterest"])
        list.append([name, date ,time,price ,price,price,price ,volume, openint ])

        side ='PE'
        name='OPTI-'+str(inst)+'-'+str(option_chain['filtered']['data'][i][side]['strikePrice'])+side
        date =now.strftime('%Y%m%d')
        time = now.strftime('%H:%M:%S')
        price = option_chain['filtered']['data'][i][side]["lastPrice"]
        volume=abs(option_chain['filtered']['data'][i][side]["changeinOpenInterest"])
        openint=abs(option_chain['filtered']['data'][i][side]["openInterest"])
        list.append([name, date ,time,price ,price,price,price ,volume, openint ])
    df=pd.DataFrame(list)
    df.to_csv(TempFile, index=False,header=None)
    
    AmiBroker.Import(0, TempFile, "amicomopti.format")
    AmiBroker.RefreshAll()

def RT(lClose):
    #logMe("RT selected")
    
    return 0 # under dev, need to use yahoo-live to fetch ticks using webhooks
    # path =TempFile
    # open(path, 'w').close()
    # global lastClose
    # continous = 0

    inst = AmiBroker.ActiveDocument.Name
    #Some how get data from yliveticker and then do the following

    # asking_time = prices[count].get("time")

    # asking_time = asking_time.replace("-", "")

    # asking_hhmm = asking_time[9:14]
    # asking_time = asking_time[:8]
    # asking_time_MST = asking_time
    # #datetimeobject = datetime.datetime.strptime(asking_time, '%d/%m/%Y %H:%M:%S')
    for t in range(1613191210,1613191210+360,10):
        
        asking_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(t))
        # asking_open = prices[count].get("openMid")
        asking_open = 200+(random.randint(0,10)) #prices[count - 1].get("closeMid")
        # asking_low = 8 #prices[count].get("lowMid")
        # asking_high = 12 # prices[count].get("highMid")
        # asking_close = 11 #prices[count].get("closeMid")
        # # if lClose != asking_close:
        ticker = AmiBroker.Stocks.Add(inst)
        quote = ticker.Quotations.Add(asking_time)
        # print(asking_time+' '+asking_hhmm)
        quote.Open = asking_open
        quote.Low = asking_open
        quote.High = asking_open
        quote.Close = asking_open
        #print(inst+str(t)+" "+str(asking_open))
        AmiBroker.RefreshAll()
        #lastClose = asking_close

        # print(asking_time,asking_open,asking_close)
        #AmiBroker.RefreshAll()

def BNBBackfill():
     #return 0
    #global daysToFil
    path = TempFile
    open(path, 'w').close()
    #file = open(path, 'w')
    Qty = AmiBroker.Stocks.Count
    
    days2Fill = int(daystofill.get()) 

    # s_date = datetime.datetime.now()-datetime.timedelta(days = int(days2Fill))
    # e_date =  datetime.datetime.now()+datetime.timedelta(days = 1)

    # start_date = s_date.strftime("%d %b, %Y")
    # end_date = e_date.strftime("%d %b, %Y")

    for i in range(0, Qty):
        inst = AmiBroker.Stocks(i).Ticker        
        if(IsOption(inst)):
            setOptions(opti2inst(inst))
            continue
        
        #logging.debug("Getting data for "+str(inst))
        tickerData = Convert2('n',inst) #remove .bnb
        #print(tickerData)
        try:
            #klines = Bclient.get_historical_klines(tickerData, Client.KLINE_INTERVAL_30MINUTE, start_date, end_date)
            klines = Bclient.get_historical_klines(tickerData, Client.KLINE_INTERVAL_30MINUTE, str(days2Fill)+" day ago UTC")
            logMe(str(len(klines)))
            df = pd.DataFrame(klines)
            #print(df.head(1))
            #logging.debug("Got data for "+str(inst))
            timelist = [datetime.datetime.fromtimestamp(x/1000)for x in df[0]]
            ticker=[inst]*len(klines)
            ticker =[Convert2('n',x) for x in ticker]
            ymd = [ x.strftime('%Y%m%d') for x in timelist  ]
            time =  [ x.strftime('%H:%M') for x in timelist  ]
            asking_open = df[1]
            asking_low = df[3]
            asking_high = df[2]
            asking_close = df[4]
            asking_volume = df[5]
            d = [ticker,ymd,time,asking_open,asking_high,asking_low,asking_close,asking_volume ]
            dfa = pd.DataFrame(data=d).transpose()
            dfa.to_csv(path, index=False,header=None)
            AmiBroker.Import(0, path, "amicom.format")
            AmiBroker.RefreshAll()
        except:
            logMe('coudnt fill '+ tickerData)

def BNBRefresh():
         #return 0
    #global daysToFil
    path = TempFile
    open(path, 'w').close()
    #file = open(path, 'w')
    Qty = AmiBroker.Stocks.Count
    
    days2Fill = int(daystofill.get()) 

    # s_date = datetime.datetime.now()-datetime.timedelta(days = int(days2Fill))
    # e_date =  datetime.datetime.now()+datetime.timedelta(days = 1)

    # start_date = s_date.strftime("%d %b, %Y")
    # end_date = e_date.strftime("%d %b, %Y")

    for i in range(0, Qty):
        inst = AmiBroker.Stocks(i).Ticker        
        if(IsOption(inst)):
            setOptions(opti2inst(inst))
            continue
        
        #logging.debug("Getting data for "+str(inst))
        tickerData = Convert2('n',inst) #remove .bnb
        #print(tickerData)
        try:
            #klines = Bclient.get_historical_klines(tickerData, Client.KLINE_INTERVAL_30MINUTE, start_date, end_date)
            klines = Bclient.get_historical_klines(tickerData, Client.KLINE_INTERVAL_30MINUTE, "1 day ago UTC")
            logMe(str(len(klines)))
            df = pd.DataFrame(klines)
            #print(df.head(1))
            #logging.debug("Got data for "+str(inst))
            timelist = [datetime.datetime.fromtimestamp(x/1000)for x in df[0]]
            ticker=[inst]*len(klines)
            ticker =[Convert2('n',x) for x in ticker]
            ymd = [ x.strftime('%Y%m%d') for x in timelist  ]
            time =  [ x.strftime('%H:%M') for x in timelist  ]
            asking_open = df[1]
            asking_low = df[3]
            asking_high = df[2]
            asking_close = df[4]
            asking_volume = df[5]
            d = [ticker,ymd,time,asking_open,asking_high,asking_low,asking_close,asking_volume ]
            dfa = pd.DataFrame(data=d).transpose()
            dfa.to_csv(path, index=False,header=None)
            AmiBroker.Import(0, path, "amicom.format")
            AmiBroker.RefreshAll()
        except:
            logMe('coudnt fill '+ tickerData)
    

def CloseAmi():
    AmiBroker.RefreshAll()
    AmiBroker.SaveDatabase()
    if tkMessageBox.askokcancel("Quit", "You want to quit now?"):
        top.destroy()

def logMe(msg):
    logging.warning(msg)
    log = (datetime.datetime.now().strftime('%d-%b-%y %H:%M:%S')+' '+msg+'\n')
    text_area.insert(tkinter.INSERT,log)
    text_area.see('end')




# Main...        
top = tkinter.Tk()
top.title("AmiCOM")
top.protocol("WM_DELETE_WINDOW", CloseAmi)

L1 = tkinter.Label(top, text=" DB Settings")
L1.pack()

L2 = tkinter.Label(top, text=" Choose DB:")
L2.pack()

DB= tkinter.StringVar(top) # choose DB
DB.set("BNBDB")
DBMenu = tkinter.OptionMenu(top, DB,"NIFTY50" ,"NIFTY100", "NIFTY200", "CUSTOM1","NEAREXP","BNBDB")
DBMenu.pack()

L3 = tkinter.Label(top, text="Days to backfill \n (max 60 for 5min and 7 for 1min)")
L3.pack()
daystofill = tkinter.StringVar()
daystofill.set(6)
E = tkinter.Entry(top, textvariable=daystofill)
E.pack()


B0 = tkinter.Button(top, text="Import tickers", command=ImportTickers)
B0.pack()
B1 = tkinter.Button(top, text="Backfill all", command=ImportThreaded)
B1.pack()
B2 = tkinter.Button(top, text="Backfill current", command=ImportCur)
B2.pack()

L4 = tkinter.Label(top, text="Auto Update Settings")
L4.pack()

isupdate = tkinter.IntVar() # Auto update or not
isupdate.set(0)
C0 = tkinter.Checkbutton(top, text="Auto Update DB", variable=isupdate, \
                 onvalue=1, offvalue=0, \
                 width=20)

C0.pack()


L5 = tkinter.Label(top, text="Update Frequenc:")
L5.pack()

refreshrate = tkinter.StringVar(top) #refresh rate 2 min 5min or 1hr
refreshrate.set("2min") # default value
refreshrateMenu = tkinter.OptionMenu(top, refreshrate,"30sec" ,"2min", "5min", "1hr")
refreshrateMenu.pack()


isRT = tkinter.IntVar() # realtime or not
isRT.set(0)

C1 = tkinter.Checkbutton(top, text="Real time (Only Current)", variable=isRT, \
                 onvalue=1, offvalue=0, \
                 width=20)

C1.pack()



B3 = tkinter.Button(top, text="Exit", command=CloseAmi)
B3.pack()
L5 = tkinter.Label(top, text="Logs:")
L5.pack()
text_area = st.ScrolledText(top,width = 40,height = 8,font = ("Times New Roman",10)) 
text_area.pack() 


nextfill = time.time() 

currentDB = "BNBDB"
while True:
    if isRT.get() == 1:
        RT(lastClose)
    daysToFill = daystofill.get()
    if (datetime.datetime.now().hour >= 9 and datetime.datetime.now().hour < 16 and isupdate.get()==1 ):
        
        if(refreshrate.get()=="30sec" and time.time()>nextfill): ## Check if db needs update
            logMe("Updating selected DB")
            nextfill = time.time()+30
            QuickImportThreaded()
            refreshOPtions()

        elif(refreshrate.get()=="2min" and time.time()>nextfill):
            logMe("Updating selected DB")
            nextfill = time.time()+2*60
            QuickImportThreaded()
            refreshOPtions()

        elif(refreshrate.get()=="5min" and time.time()>nextfill):
            logMe("Updating selected DB")
            nextfill = time.time()+5*60
            QuickImportThreaded()
            refreshOPtions()

        elif(refreshrate.get()=="2min" and time.time()>nextfill):
            logMe("Updating selected DB")
            nextfill = time.time()+60*60
            QuickImportThreaded()
            refreshOPtions()


    if(currentDB!=DB.get()):  ### Check if DB has changed
        if(DB.get()=="NIFTY50"):
            AmiBroker.LoadDatabase(NIFTY50DB)
        elif(DB.get()=="NIFTY100"):
            AmiBroker.LoadDatabase(NIFTY100DB)
        elif(DB.get()=="NIFTY200"):
            AmiBroker.LoadDatabase(NIFTY200DB)
        elif(DB.get()=="CUSTOM1"):
            AmiBroker.LoadDatabase(CUSTOM1DB)
        elif(DB.get()=="NEAREXP"):
            AmiBroker.LoadDatabase(NEAREXPDB)
        elif(DB.get()=="BNBDB"):
            AmiBroker.LoadDatabase(BNBDB)
        currentDB = DB.get()           
    top.update_idletasks()
    top.update()
    time.sleep(0.0001)

