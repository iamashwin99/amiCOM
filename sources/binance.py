from exchange import exchange


class Binance(exchange):
    def __init__(self):
        loadedDBs['binance'] = {
            "dbaselocation":"",
            "dblistlocation":""
        }
