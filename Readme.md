# AmiCOM

AmiCOM (development in progress ) is a python utility to automatically import 

Data from Yahoo finance into Amibroker.



# How does it work

AmiCOM downloads the required data from yahoo finance for the specified script, time frame and available resolution (more on this later) and pushes it onto a temp file, through OLE automation it then imports them into the specified database. 

Knowing the location of database is important and thus AmiCOM makes use of 3 pre programmed databases (and stored in DB.zip). The 3 pre-programmed dbs include NIFTY50,NIFTY100,NIFTY200 stocks respectively including the major indices . 

It is important to note that yahoo finance allows at most 2000 api calls per hour per ip. Thus you need to ensure that you dont cross that limit by managing the number of scripts being refreshed and the refresh rate. eg in the Nifty50 db there are about 60 stocks so 2000 queries per hr/60 scripts =  33.33 database refreshes every hour or  about 1 complete query every 2 mins. So if you are using Nifty50 db you are suggested to refresh not more frequently than once every 2 mins, however you can easily choose higher refresh rate ( like once every 15 mins) without any trouble.



## Data availability



Yahoo finance provides selective availability to intraday data

* for one minute candle only  past 7 days of data is available
* for 5 min to 1hr candles only past 60 days of data is available
* 1D data is available for mulitple years

AmiCOM automatically chooses the best candle format for the given back fill time frame.

If you choose to back fill 6 days of tick then they will be filled in 1min candles

If you choose to back fill up-to past 60 days then 5 min candles will be chosen

Anything beyond that only EOD data is fetched. 



# How to use this?

\* Clone this repository into the root of your C drive (VERY IMPORTANT).

\* Unzip DB.zip and TickerList.zip within this folder

\* install required modules (look at main.py)

\* Copy amicom.format into the format ofler inside Amibroker

\* run `python main.py`

\* Enjoy!

If you need to add a new symbol use amibroker to make a new one and then in AmiCOM click on Back fill current



![snap.png](snap.png)


Please not that this project is in active development and must not be used in any  serious work, use this at your own discretion.  


Collaborators are welcomed to send PRs and post issues on Github. 



AmiCOM is provided with GPL license thus if you use parts of this project in your project you are required to release its source code.