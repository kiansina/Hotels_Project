import pandas as pd
from docx import Document
from docx.shared import Inches,Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.shared import RGBColor
from docx.enum.table import WD_ROW_HEIGHT_RULE
import matplotlib.pyplot as plt
import math
from matplotlib.figure import Figure
import numpy as np
import os
import matplotlib
import locale
#matplotlib.use('Agg')
locale.setlocale(locale.LC_ALL, 'it_IT')
matplotlib.style.use('seaborn')
df=pd.read_excel(r"C:\Users\sina.kian\Desktop\Ricardo\DataBase_3_final.xlsx")
dp=pd.read_excel(r"C:\Users\sina.kian\Desktop\Ricardo\updates\Fina_Merge.xlsx")

df=df[df.columns[1:]]
df['Codice_Unico']=range(0,len(df))
for i in range(0,len(df)):
    df['Codice_Unico'][i]=str(df['Codice Hotel'][i])+'_'+str(df['Denominazione Hotel'][i])

############################################
############################################
##PREPARING NEW DATA ARRIVED:
############################################
############################################
dp['Codice_Unico']=range(0,len(dp))
for i in range(0,len(dp)):
    dp['Codice_Unico'][i]=str(dp['Codice Hotel'][i])+'_'+str(dp['Denominazione hotel'][i])


ll=[]
LL=[]
for i in dp['Codice_Unico']:
    ll=df[df['Codice_Unico']==i].index.to_list()
    LL+=ll

#################### CHECK to be True
df.loc[LL]['Codice_Unico'].to_list()==dp['Codice_Unico'].to_list()
#################### CHECK to be True

dp.columns=['Denominazione Hotel', 'Codice Hotel',
       'Fatturato  Hotel 2019 (€)', 'Fatturato  Hotel 2020 (€)',
              'Fatturato  Hotel 2021 (€)', 'Fatturato  Hotel 2022 (€)',
       'Codice_Unico']


dp.set_axis(LL,inplace=True)
df.update(dp)

dp.to_excel("Fina_Merge_LL.xlsx")
df.to_excel('DataBase_4_final.xlsx')
