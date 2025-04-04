 
import win32com.client 
import sys
import re
import glob
import csv
import shutil
import os
import pandas

from datetime import datetime as dt
from pdd_utils import * 

import configparser

'''
Ekstra raport z PDD wind0 dla wszystkich Agentów o ile Total polisy >0
Dla wyjątkowych wszystkie cancelled niezależnie od Total Policy 
Data utworzenia polisy w wybranym zakresie
Status cancelled
Kolumny takie jak w zestawianiu płatności zbiorczych polisa;termin;agent;klient;data wystawienia polisy;kwota;VIN;OFWCA

'''

MAIN_DIR    = os.path.dirname(os.path.abspath(sys.argv[0])) 


version = "1.0"
title = 'pdd_anulowane_zbiorcze'
title_ok = 'pdd_anulowane_zbiorcze'
script_name = "pdd_anulowane_zbiorcze"



def main():

    parser = configparser.ConfigParser()
    #print(os.path.join(MAIN_DIR, 'config.txt'))
    parser.read(os.path.join(MAIN_DIR, 'config.txt'),  encoding='utf-8')
    AGENCJE_SPECJALNE      = parser.get("config", "AGENCJE_SPECJALNE")
    start_date      = parser.get("config", "start_date")
    stop_date      = parser.get("config", "stop_date")

    special = AGENCJE_SPECJALNE.split(',')

    df = get_df_newest_pdd()

    df = df[(df['Policy issue date']>=start_date) & (df['Policy issue date']<=stop_date) & (df['Status'] == 'cancelled')]
    df_out = df[(df['Total policy annual premium [pln]'].astype(float) > 0)]

    for agent in special:
        df1 = df[(df['Total policy annual premium [pln]'].astype(float) <= 0) & (df['Agent company name'] == agent)]
        df_out = pandas.concat([df_out, df1])

    df_out = df_out[['Policy number', 'Agent company name','Policy holder first Name','Policy holder last name','Policy holder company name','Policy issue date','Total policy annual premium [pln]', 'VIN No.','Agent natural person name and last name']]

    if os.path.isfile(MAIN_DIR + "\\pdd_cancelled.xlsx"):
        os.remove(MAIN_DIR + "\\pdd_cancelled.xlsx")
        
    df_out.to_excel(MAIN_DIR + "\\pdd_cancelled.xlsx", index=False)


    

                            



if __name__ == "__main__":
    main()



    