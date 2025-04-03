 
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

'''
Ekstra raport z PDD wind0 dla wszystkich Agentów o ile Total polisy >0
Dla wyjątkowych wszystkie cancelled niezależnie od Total Policy 
Data utworzenia polisy w wybranym zakresie
Status cancelled
Kolumny takie jak w zestawianiu płatności zbiorczych polisa;termin;agent;klient;data wystawienia polisy;kwota;VIN;OFWCA

'''

special = ['MUBI']

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.append(os.path.dirname(CURRENT_DIR))

version = "1.0"
title = 'pdd_anulowane_zbiorcze'
title_ok = 'pdd_anulowane_zbiorcze'
script_name = "pdd_anulowane_zbiorcze"



def main():
    df = get_df_newest_pdd()

    start_date = '2020-01-01'
    stop_date = '2025-11-11'
    df = df[(df['Policy issue date']>=start_date) & (df['Policy issue date']<=stop_date) & (df['Status'] == 'cancelled')]

    df_out = df[(df['Total policy annual premium [pln]'].astype(float) > 0)]

    for agent in special:
        df1 = df[(df['Total policy annual premium [pln]'].astype(float) <= 0) & (df['Agent company name'] == agent)]
        df_out = pandas.concat([df_out, df1])

    df_out = df_out[['Policy number', 'Agent company name','Policy holder first Name','Policy holder last name','Policy holder company name','Policy issue date','Total policy annual premium [pln]', 'VIN No.','Agent natural person name and last name']]
    df_out.to_excel("pdd_cancelled.xlsx", index=False)


    

                            



if __name__ == "__main__":
    main()



    