
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



def Backfill():
    #return 0


    days2Fill = int(df.get()) 
    if days2Fill < 7:
        interval_length = '1m'
    elif days2Fill < 60:
        interval_length = '5m'
    else:
        interval_length = '1d'

    s_date = datetime.datetime.now()-datetime.timedelta(days = int(days2Fill))
    e_date =  datetime.datetime.now()

    start_date = s_date.strftime("%Y-%m-%d")
    end_date = e_date.strftime("%Y-%m-%d")

    continous = 0
    while continous == 0:
        Qty = AmiBroker.Stocks.Count
        for i in range(0, Qty):
            inst = AmiBroker.Stocks(i).Ticker
            print("Getting data for "+str(inst))
            tickerData = yf.Ticker(inst)
            tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
            print("Got data for "+str(inst))
            timelist = list(tickerDf.index)
            #response = oanda.get_history(instrument=inst, count="5000", granularity="D", candleFormat="midpoint")
            #prices = response.get("candles")
            for count in range(0, len(tickerDf)):                
                asking_time = timelist[count].strftime('%d/%m/%Y %H:%M:%S')
                asking_open = tickerDf['Open'][count]
                asking_low = tickerDf['Low'][count]
                asking_high = tickerDf['High'][count]
                asking_close = tickerDf['Close'][count]
                asking_volume = tickerDf['Volume'][count]                
                ticker = AmiBroker.Stocks.Add(inst)
                quote = ticker.Quotations.Add(asking_time)
                #print("Adding time "+str(asking_time))
                quote.Open = asking_open
                quote.Low = asking_low
                quote.High = asking_high
                quote.Close = asking_close
                AmiBroker.RefreshAll()
                #print(asking_time,asking_open,asking_close)


def Import():
    #return 0
    #global daysToFil
    path = 'C:\\amiCOM\\temp.txt'
    open(path, 'w').close()
    #file = open(path, 'w')
    Qty = AmiBroker.Stocks.Count
    
    days2Fill = int(df.get()) 
    if days2Fill < 7:
        interval_length = '1m'
    elif days2Fill < 60:
        interval_length = '5m'
    else:
        interval_length = '1d'

    s_date = datetime.datetime.now()-datetime.timedelta(days = int(days2Fill))
    e_date =  datetime.datetime.now()

    start_date = s_date.strftime("%Y-%m-%d")
    end_date = e_date.strftime("%Y-%m-%d")

    for i in range(0, Qty):
        inst = AmiBroker.Stocks(i).Ticker
        print("Getting data for "+str(inst))
        tickerData = yf.Ticker(inst)
        tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
        print("Got data for "+str(inst))
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
    open(path, 'w').close()
    path = 'C:\\amiCOM\\temp.txt'
    days2Fill = int(df.get()) 
    if days2Fill < 7:
        interval_length = '1m'
    elif days2Fill < 60:
        interval_length = '5m'
    else:
        interval_length = '1d'

    s_date = datetime.datetime.now()-datetime.timedelta(days = int(days2Fill))
    e_date =  datetime.datetime.now()

    start_date = s_date.strftime("%Y-%m-%d")
    end_date = e_date.strftime("%Y-%m-%d")

    inst = AmiBroker.ActiveDocument.Name
    print("Getting data for "+str(inst))
    tickerData = yf.Ticker(inst)
    tickerDf = tickerData.history(interval=interval_length, start=start_date, end=end_date)
    print("Got data for "+str(inst))
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
    open(path, 'w').close()
    global lastClose
    continous = 0

    inst = AmiBroker.ActiveDocument.Name
    response = oanda.get_history(instrument=inst, count="2", granularity="D", candleFormat="midpoint")
    prices = response.get("candles")

    for count in range(1, len(prices)):
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
