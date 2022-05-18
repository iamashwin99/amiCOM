import datetime  # For datetime objects
import os  # To manage paths
import sys
from typing import NoReturn
from win32com.client import Dispatch

from sys import argv
import tkinter
import time
import math
import pandas as pd

import logging
import re
import tkinter.messagebox as tkMessageBox
# ttk makes the window look like running Operating System’s theme
from tkinter import ttk
import tkinter.scrolledtext as st 
import random

from decouple import config
loadedDBs={
    "exchange":{
        "dbaselocation":"",
        "dblistlocation":""
    }
}

open(LOGDIR, 'w').close() # Clear log file while first load
logging.basicConfig(format='%(asctime)s - %(message)s', datefmt='%d-%b-%y %H:%M:%S')
#logging.basicConfig(filename=LOGDIR, level=logging.debug, format='%(asctime)s - %(levelname)s - %(message)s')

#Open Amibroker with COM32 conection
AmiBroker = Dispatch("Broker.Application")
AmiBroker.visible=True

#AmiBroker.LoadDatabase(BNBDB)




def CloseAmi():
    refreshAmi()
    if tkMessageBox.askokcancel("Quit", "You want to quit now?"):
        top.destroy()
def refreshAmi():
    AmiBroker.RefreshAll()
    AmiBroker.SaveDatabase()

def logMe(msg):
    logging.warning(msg)
    log = (datetime.datetime.now().strftime('%d-%b-%y %H:%M:%S')+' '+msg+'\n')
    text_area.insert(tkinter.INSERT,log)
    text_area.see('end')
    top.update()




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
DBMenu = tkinter.OptionMenu(top, DB,"NIFTY50" ,"NIFTY100", "NIFTY200", "CUSTOM1","NEAREXP","BNBDB","KCSDB")
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
refreshrate.set("1hr") # default value
refreshrateMenu = tkinter.OptionMenu(top, refreshrate,"30sec" ,"2min", "5min", "1hr")
refreshrateMenu.pack()



isRT = tkinter.IntVar() # realtime or not
isRT.set(0)

C1 = tkinter.Checkbutton(top, text="Real time (Only Options)", variable=isRT, \
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
nextRT = time.time() 

currentDB = "BNBDB"
while True:
    if (isRT.get() == 1 and time.time()>nextRT):
        RT(lastClose)
        nextRT  = time.time() + 1

    daysToFill = daystofill.get()
    if ( isupdate.get()==1 ):
        # if not (  DB.get()=="BNBDB" or DB.get()=="KCSDB" ): #we are pulling nse data
        #     if not (( (datetime.datetime.now().hour >= 9 and datetime.datetime.now().hour < 16)):
        #         break
        
        if(refreshrate.get()=="30sec" and time.time()>nextfill): ## Check if db needs update
            logMe("Updating selected DB")
            nextfill = time.time()+30
            QuickImportThreaded()
            

        elif(refreshrate.get()=="2min" and time.time()>nextfill):
            logMe("Updating selected DB")
            nextfill = time.time()+2*60
            QuickImportThreaded()
            

        elif(refreshrate.get()=="5min" and time.time()>nextfill):
            logMe("Updating selected DB")
            nextfill = time.time()+5*60
            QuickImportThreaded()
            

        elif(refreshrate.get()=="1hr" and time.time()>nextfill):
            logMe("Updating selected DB")
            nextfill = time.time()+60*60
            QuickImportThreaded()
            


    if(currentDB!=DB.get()):  ### Check if DB has changed
        if(DB.get()=="NIFTY50"):
            refreshAmi()
            AmiBroker.LoadDatabase(NIFTY50DB)
        elif(DB.get()=="NIFTY100"):
            refreshAmi()
            AmiBroker.LoadDatabase(NIFTY100DB)

        elif(DB.get()=="NIFTY200"):
            refreshAmi()
            AmiBroker.LoadDatabase(NIFTY200DB)
        elif(DB.get()=="CUSTOM1"):
            refreshAmi()
            AmiBroker.LoadDatabase(CUSTOM1DB)
        elif(DB.get()=="NEAREXP"):
            refreshAmi()
            AmiBroker.LoadDatabase(NEAREXPDB)
        elif(DB.get()=="BNBDB"):
            refreshAmi()
            AmiBroker.LoadDatabase(BNBDB)
        elif(DB.get()=="KCSDB"):
            refreshAmi()
            AmiBroker.LoadDatabase(KCSDB)
        currentDB = DB.get()           
    top.update_idletasks()
    top.update()
    time.sleep(0.001)
