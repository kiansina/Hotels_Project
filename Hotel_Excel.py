import pandas as pd

df=pd.read_excel(r"*****")
df.columns=df.loc[5]
df=df[6:168].set_index(i for i in range(0,len(df[6:168])))
df=df.loc[:,~df.columns.duplicated()]
df['Franchigia Danni Indiretti']=['5 Giorni']*len(df)
df['Fatturato totale']=[0]*len(df)
for i in range(0,len(df)):
    df['Fatturato totale'][i]=df['Fatturato  HOTEL 2019'][i]+df['Fatturato Ristorante 2019'][i]

NN=['Franchigia Danni Diretti',
       'Valore fabbricato', 'Valore contenuti', 'Ricorso Terzi', 'Cristalli',
       'Furto', 'Merci Refrigerazione', 'Fenomeno Elettrico', 'Margine di contribuzione', 'TOTALE FABBRICATO + CONTENUTO']

#NN=['Franchigia Danni Diretti',
#       'Valore fabbricato', 'Valore contenuti', 'Ricorso Terzi', 'Cristalli',
#       'Furto', 'Merci Refrigerazione', 'Fenomeno Elettrico', 'Margine di contribuzione', 'TOTALE FABBRICATO + CONTENUTO', 'Premio Imponibile Annuo',
#       'Premio Imponibile PRORATA', 'Premio Lordo Annuo',
#       'Premio Lordo PRORATA', 'Franchigia (RCT)', 'Fatturato Ristorante 2019',
#       'Fatturato Ristorante 2020', 'Fatturato  HOTEL 2019',
#       'Fatturato HOTEL 2020', 'RC Terzi (RCT)', 'RC Prodotti (RCP) e Smercio',
#       'RC Prestatori di Lavoro (RCO)',
#       'RC Prestatori di Lavoro (RCO) - per persona','RC Terzi (RCT) Cose non consegnate / Garage keeper\'s liability',
#       'Premio lordo Annuo 2019', 'Premio Lordo Annuo 2020', 'Fatturato totale']
for i in NN:
    df[i].fillna(0,inplace=True)

for i in NN:
        df[i]=df[i].map('{:.d}'.format)
        #for j in range(0,len(df)):
            #df[i].loc[j]=round(float(df[i].loc[j]))
            #df[i]='{:,}'.format(df[i]).replace(',','.')




dff=df[7:8]
#######
dT1=dff[['Codice Hotel','Assicurato','Indirizzo (Sede Legale)','Città (Sede Legale)','Provincia (Sede Legale)','Cap\n(Sede Legale)','C.F./P.IVA']]

dT2=dff[['Codice Hotel','Denominazione Hotel','Indirizzo (Ubicazione Hotel)','Città (Ubicazione Hotel)','Provincia \n(Ubicazione Hotel)','Cap (Ubicazione Hotel)']]

dT3=dff[['Codice Hotel','Valore fabbricato', 'Valore contenuti', 'Ricorso Terzi','Terremoto, Inondazione, Alluvione, Allagamento',
       'Terrorismo, Eventi Socio Politici, Atti Dolosi',
       'Margine di contribuzione', 'Periodo di Indennizzo (Mesi)',
       'Cimici da letto\n(se 0 o "-" non operante; numero mesi se operante)']]

dT4=dff[['Codice Hotel', 'Franchigia Danni Diretti', 'Franchigia Danni Indiretti', 'Zona rischi catastrofali Nr.']]

dT5=dff[['Codice Hotel', 'Cristalli', 'Furto', 'Merci Refrigerazione', 'Fenomeno Elettrico']]

dT6=dff[['Codice Hotel', 'Fatturato  HOTEL 2019', 'Fatturato Ristorante 2019', 'Franchigia (RCT)']]

dT7=dff[['Codice Hotel', 'Presenza di vincolo']]

for i in dff['Codice Hotel']:
    di1=dT1[dT1['Codice Hotel']==i]
    di1=di1.T
    di1=di1.reset_index()
    di1.columns=di1.loc[0]
    di1=di1.iloc[1:]
    di1.set_index('Codice Hotel',inplace=True)
    #
    di2=dT2[dT2['Codice Hotel']==i]
    di2=di2.T
    di2=di2.reset_index()
    di2.columns=di2.loc[0]
    di2=di2.iloc[1:]
    di2.set_index('Codice Hotel',inplace=True)
    #
    di3=dT3[dT3['Codice Hotel']==i]
    di3=di3.T
    di3=di3.reset_index()
    di3.columns=di3.loc[0]
    di3=di3.iloc[1:]
    di3.set_index('Codice Hotel',inplace=True)
    #
    di4=dT4[dT4['Codice Hotel']==i]
    di4=di4.T
    di4=di4.reset_index()
    di4.columns=di4.loc[0]
    di4=di4.iloc[1:]
    di4.set_index('Codice Hotel',inplace=True)
    #
    di5=dT5[dT5['Codice Hotel']==i]
    di5=di5.T
    di5=di5.reset_index()
    di5.columns=di5.loc[0]
    di5=di5.iloc[1:]
    di5.set_index('Codice Hotel',inplace=True)
    #
    di6=dT6[dT6['Codice Hotel']==i]
    di6=di6.T
    di6=di6.reset_index()
    di6.columns=di6.loc[0]
    di6=di6.iloc[1:]
    di6.set_index('Codice Hotel',inplace=True)
    #
    di7=dT7[dT7['Codice Hotel']==i]
    di7=di7.T
    di7=di7.reset_index()
    di7.columns=di7.loc[0]
    di7=di7.iloc[1:]
    di7.set_index('Codice Hotel',inplace=True)
    #
    name=str(i)+'.xlsx'
    writer = pd.ExcelWriter(name,engine='xlsxwriter')
    workbook=writer.book
    di1.to_excel(writer,sheet_name='Information',startrow=1 , startcol=1)
    di2.to_excel(writer,sheet_name='Information',startrow=12, startcol=1)
    di3.to_excel(writer,sheet_name='Information',startrow=22, startcol=1)
    di4.to_excel(writer,sheet_name='Information',startrow=35, startcol=1)
    di5.to_excel(writer,sheet_name='Information',startrow=43, startcol=1)
    di6.to_excel(writer,sheet_name='Information',startrow=52, startcol=1)
    di7.to_excel(writer,sheet_name='Information',startrow=60, startcol=1)
    worksheet='Information'
    # Auto-adjust columns' width
    #for column in df:
    #column_width = max(df[column].astype(str).map(len).max(), len(column))
    #col_idx = df.columns.get_loc(column)
    #writer.sheets['Information'].set_column(col_idx, col_idx, column_width)
    # Manually adjust the wifth of column 'this_is_a_long_column_name'
    format = workbook.add_format()
    format.set_align('center')
    format.set_align('vcenter')
    col_idx = df.columns.get_loc('Codice Hotel')
    writer.sheets['Information'].set_column(col_idx, col_idx, 59.29,format)
    col_idx = di1.columns.get_loc(di1.columns[0])
    writer.sheets['Information'].set_column(col_idx, col_idx, 42.43,format)
    writer.sheets['Information'].set_column('C:C', 42.43,format)
    #
    writer.save()

   # Creating Excel Writer Object from Pandas

#df.to_excel
