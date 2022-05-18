from exchange import exchange
import yahoo_finance as yf
import pandas as pd

class NSE(exchange):
    def __init__(self):
        loadedDBs['nse'] = {
            "dbaselocation":"",
            "dblistlocation":""
        }
