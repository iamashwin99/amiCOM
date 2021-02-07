import datetime  # For datetime objects
import os  # To manage paths
import sys
from win32com.client import Dispatch
import datetime 
from sys import argv
import tkinter
import time
import pandas as pd
import yfinance as yf
import logging
import tkinter.messagebox as tkMessageBox
# ttk makes the window look like running Operating System’s theme
from tkinter import ttk
import tkinter.scrolledtext as st 

lastClose = 0
abDatabase = 'C:\\amiCOM\\DB'
NIFTY50DB = 'C:\\amiCOM\\DB\\NIFTY50'
NIFTY100DB = 'C:\\amiCOM\\DB\\NIFTY100'
NIFTY200DB = 'C:\\amiCOM\\DB\\NIFTY200'
CUSTOM1DB = 'C:\\amiCOM\\DB\\CUSTOM1'

TempFile= 'C:\\amiCOM\\temp.txt'
open(TempFile, 'w').close() # Clear temp file while first load

NIFTY50list = 'C:\\amiCOM\\TickerList\\NIFTY50.txt'
NIFTY100List = 'C:\\amiCOM\\TickerList\\NIFTY100.txt'
NIFTY200List = 'C:\\amiCOM\\TickerList\\NIFTY200.txt'
CUSTOM1List = 'C:\\amiCOM\\TickerList\\CUSTOM1.txt'
LOGDIR = 'C:\\amiCOM\\Logs.txt'

open(LOGDIR, 'w').close() # Clear log file while first load


logging.basicConfig(format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')
#logging.basicConfig(filename=LOGDIR, level=logging.debug, format='%(asctime)s - %(levelname)s - %(message)s')
AmiBroker = Dispatch("Broker.Application")
AmiBroker.visible=True

AmiBroker.LoadDatabase(NIFTY50DB)



## Methods

def ImportTickers():
    
    source=DB.get() #"NIFTY50" ,"NIFTY100", "NIFTY200", "CUSTOM1"
    ticker =[]
    if(source =="NIFTY50" ): 
        filename = NIFTY50list

    elif(source =="NIFTY100" ): 
        filename = NIFTY100List

    elif(source =="NIFTY200" ): 
        filename = NIFTY200List
    else: 
        filename = CUSTOM1List

    with open(filename) as f:
        ticker = f.readlines()
        ticker = [x.strip() for x in ticker] # remove whitespace characters like `\n` at the end of each line
    for count in range(0, len(ticker)):
        AmiBroker.Stocks.Add(ticker[count])
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
            tickerData = yf.Ticker(inst)
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
                ticker = AmiBroker.Stocks.Add(inst)
                quote = ticker.Quotations.Add(asking_time)
                quote.Open = asking_open
                quote.Low = asking_low
                quote.High = asking_high
                quote.Close = asking_close
                AmiBroker.RefreshAll()

def Importthreaded():
     
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
        listofstocks.append(AmiBroker.Stocks(i).Ticker)
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
    

def Import():
    
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

    for i in range(0, Qty):
        inst = AmiBroker.Stocks(i).Ticker
        #logging.debug("Getting data for "+str(inst))
        tickerData = yf.Ticker(inst)
        tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
        #logging.debug("Got data for "+str(inst))
        timelist = list(tickerDf.index)
        ticker=[inst]*len(tickerDf)
        ymd = [ x.strftime('%Y%m%d') for x in timelist  ]
        time =  [ x.strftime('%H:%M') for x in timelist  ]
        asking_open = tickerDf['Open']
        asking_low = tickerDf['Low']
        asking_high = tickerDf['High']
        asking_close = tickerDf['Close']
        asking_volume = tickerDf['Volume']
        d = [ticker,ymd,time,asking_open,asking_high,asking_low,asking_close,asking_volume ]
        dfa = pd.DataFrame(data=d).transpose()
        dfa.to_csv(path, index=False,header=None)
        AmiBroker.Import(0, path, "amicom.format")
        AmiBroker.RefreshAll()
    



def QuickImport():
    
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

    for i in range(0, Qty):
        inst = AmiBroker.Stocks(i).Ticker
        #logging.debug("Getting data for "+str(inst))
        tickerData = yf.Ticker(inst)
        tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
        #logging.debug("Got data for "+str(inst))
        timelist = list(tickerDf.index)
        ticker=[inst]*len(tickerDf)
        ymd = [ x.strftime('%Y%m%d') for x in timelist  ]
        time =  [ x.strftime('%H:%M') for x in timelist  ]
        asking_open = tickerDf['Open']
        asking_low = tickerDf['Low']
        asking_high = tickerDf['High']
        asking_close = tickerDf['Close']
        asking_volume = tickerDf['Volume']
        d = [ticker,ymd,time,asking_open,asking_high,asking_low,asking_close,asking_volume ]
        dfa = pd.DataFrame(data=d).transpose()
        dfa.to_csv(path, index=False,header=None)
        AmiBroker.Import(0, path, "amicom.format")
        AmiBroker.RefreshAll()
    


def ImportCur():
    
    path =TempFile
    open(path, 'w').close()

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

    inst = AmiBroker.ActiveDocument.Name
    logMe("Getting data for "+str(inst))
    tickerData = yf.Ticker(inst)
    tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
    logMe("Got data for "+str(inst))
    timelist = list(tickerDf.index)
    ticker=[inst]*len(tickerDf)
    ymd = [ x.strftime('%Y%m%d') for x in timelist  ]
    time =  [ x.strftime('%H:%M') for x in timelist  ]
    asking_open = tickerDf['Open']
    asking_low = tickerDf['Low']
    asking_high = tickerDf['High']
    asking_close = tickerDf['Close']
    asking_volume = tickerDf['Volume']
    d = [ticker,ymd,time,asking_open,asking_high,asking_low,asking_close,asking_volume ]
    dfa = pd.DataFrame(data=d).transpose()
    dfa.to_csv(path, index=False,header=None)
    AmiBroker.Import(0, path, "amicom.format")
    AmiBroker.RefreshAll()
    


def RT(lClose):
    #logMe("RT selected")
    
    return 0 # under dev, need to use yahoo-live to fetch ticks using webhooks
    path =TempFile
    open(path, 'w').close()
    global lastClose
    continous = 0

    inst = AmiBroker.ActiveDocument.Name
    #Some how get data from yliveticker and then do the following

    asking_time = prices[count].get("time")

    asking_time = asking_time.replace("-", "")

    asking_hhmm = asking_time[9:14]
    asking_time = asking_time[:8]
    asking_time_MST = asking_time
    datetimeobject = datetime.datetime.strptime(asking_time, '%Y%m%d')
    asking_time = datetimeobject.strftime('%d/%m/%Y')
    # asking_open = prices[count].get("openMid")
    asking_open = prices[count - 1].get("closeMid")
    asking_low = prices[count].get("lowMid")
    asking_high = prices[count].get("highMid")
    asking_close = prices[count].get("closeMid")
    if lClose != asking_close:
        ticker = AmiBroker.Stocks.Add(inst)
        quote = ticker.Quotations.Add(asking_time)
        # print(asking_time+' '+asking_hhmm)
        quote.Open = asking_open
        quote.Low = asking_low
        quote.High = asking_high
        quote.Close = asking_close
        AmiBroker.RefreshAll()
        lastClose = asking_close

        # print(asking_time,asking_open,asking_close)
        # AmiBroker.RefreshAll()

    

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
DB.set("NIFTY50")
DBMenu = tkinter.OptionMenu(top, DB,"NIFTY50" ,"NIFTY100", "NIFTY200", "CUSTOM1")
DBMenu.pack()

L3 = tkinter.Label(top, text="Days to backfill \n (max 60 for 5min and 7 for 1min)")
L3.pack()
daystofill = tkinter.StringVar()
daystofill.set(6)
E = tkinter.Entry(top, textvariable=daystofill)
E.pack()


B0 = tkinter.Button(top, text="Import tickers", command=ImportTickers)
B0.pack()
B1 = tkinter.Button(top, text="Backfill all", command=Import)
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

# progress = ttk.Progressbar(top, orient = HORIZONTAL, length = 120,mode='indeterminate')
# progress.pack()
# # progress.start(10)
# # time.sleep(5)
# # progress.stop()
# scrollbar = Scrollbar(top)
# scrollbar.pack(side=RIGHT, fill=Y)
# textbox = Text(top)
# textbox.pack()
# textbox.config(yscrollcommand=scrollbar.set)
# scrollbar.config(command=textbox.yview)
L5 = tkinter.Label(top, text="Logs:")
L5.pack()
text_area = st.ScrolledText(top,width = 30,height = 8,font = ("Times New Roman",10)) 
text_area.pack() 


nextfill = time.time() 

currentDB = "NIFTY50"
while True:
    if isRT.get() == 1:
        RT(lastClose)
    daysToFill = daystofill.get()
    if (datetime.datetime.now().hour >= 9 and datetime.datetime.now().hour < 16 and isupdate.get()==1 ):
        
        if(refreshrate.get()=="30sec" and time.time()>nextfill): ## Check if db needs update
            logMe("Updating selected DB")
            nextfill = time.time()+30
            QuickImport()

        elif(refreshrate.get()=="2min" and time.time()>nextfill):
            logMe("Updating selected DB")
            nextfill = time.time()+2*60
            QuickImport()

        elif(refreshrate.get()=="5min" and time.time()>nextfill):
            logMe("Updating selected DB")
            nextfill = time.time()+5*60
            QuickImport()

        elif(refreshrate.get()=="2min" and time.time()>nextfill):
            logMe("Updating selected DB")
            nextfill = time.time()+60*60
            QuickImport()


    if(currentDB!=DB.get()):  ### Check if DB has changed
        if(DB.get()=="NIFTY50"):
            AmiBroker.LoadDatabase(NIFTY50DB)
        elif(DB.get()=="NIFTY100"):
            AmiBroker.LoadDatabase(NIFTY100DB)
        elif(DB.get()=="NIFTY200"):
            AmiBroker.LoadDatabase(NIFTY200DB)
        elif(DB.get()=="CUSTOM1"):
            AmiBroker.LoadDatabase(CUSTOM1DB)
        currentDB = DB.get()           



    top.update_idletasks()
    top.update()
    time.sleep(0.0001)

