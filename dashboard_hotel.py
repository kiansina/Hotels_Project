import dash
from dash import dcc, html
import matplotlib
import matplotlib.pyplot as plt
import plotly.graph_objs as go
import numpy as np
import pandas as pd
import random
import locale
from dash.dependencies import Input, Output


cmap = matplotlib.cm.Blues(np.linspace(0,1,40))
CCC= cmap[4]
app = dash.Dash()


locale.setlocale(locale.LC_ALL, 'it_IT')
matplotlib.style.use('seaborn')
df=pd.read_excel(r"C:\Users\sina.kian\Desktop\Ricardo\updates\DataBase_5_final.xlsx")


S=['Denominazione Hotel','Indirizzo (Ubicazione Hotel)','Provincia \n(Ubicazione Hotel)','Fabbricato',
       'Contenuto','Terremoto, Inondazione, Alluvione, Allagamento','Coef. Terremoto', 'Earthquake Zone',  'Earthquake_Risk_Score_Risk_Score', 'Flood_Risk_Score_Risk_Score','Longitude','Latitude','Codice Hotel','Codice_Unico']
#S=['Codice Hotel','Indirizzo (Ubicazione Hotel)',
#        'Provincia \n(Ubicazione Hotel)', 'Città (Ubicazione Hotel)',
#        'Cap (Ubicazione Hotel)','Terremoto, Inondazione, Alluvione, Allagamento','Coef. Terremoto','Earthquake Zone', 'Earthquake_Risk_Score_Risk_Score', 'Flood_Risk_Score_Risk_Score','Longitude','Latitude']

MP=df[['Fabbricato', 'Contenuto']]
dfs=df[S]


LOW=dfs[dfs['Earthquake Zone']==0]
LOW.sort_values(by=['Terremoto, Inondazione, Alluvione, Allagamento'], inplace=True)
MEDIUM=dfs[dfs['Earthquake Zone']==1]
MEDIUM.sort_values(by=['Terremoto, Inondazione, Alluvione, Allagamento'], inplace=True)
HIGH=dfs[dfs['Earthquake Zone']==2]
HIGH.sort_values(by=['Terremoto, Inondazione, Alluvione, Allagamento'], inplace=True)
Extream=dfs[dfs['Earthquake Zone']==3]
Extream.sort_values(by=['Terremoto, Inondazione, Alluvione, Allagamento'], inplace=True)
Disasterous=dfs[dfs['Earthquake Zone']==4]
Disasterous.sort_values(by=['Terremoto, Inondazione, Alluvione, Allagamento'], inplace=True)

app.layout = html.Div([
    html.Div([  # this Div contains our scatter plot
    dcc.Graph(
        id='mpg_scatter',
        figure={
            'data': [go.Scattergeo(
                                lon = LOW.Longitude,
                                lat = LOW.Latitude,
                                text = LOW['Codice_Unico'],
                                marker_color='green',
                                name='0',
                                marker_size=8,
                                              ),
                    go.Scattergeo(
                                        lon = MEDIUM.Longitude,
                                        lat = MEDIUM.Latitude,
                                        text = MEDIUM['Codice_Unico'],
                                        marker_color='yellow',
                                        name='1',
                                        marker_size=8,
                                                      ),
                  go.Scattergeo(
                                      lon = HIGH.Longitude,
                                      lat = HIGH.Latitude,
                                      text = HIGH['Codice_Unico'],
                                      marker_color='Orange',
                                      name='2',
                                      marker_size=8,
                                                    ),
                 go.Scattergeo(
                                     lon = Extream.Longitude,
                                     lat = Extream.Latitude,
                                     text = Extream['Codice_Unico'],
                                     marker_color='red',
                                     name='3',
                                     marker_size=8,
                                                   ),
                go.Scattergeo(
                                    lon = Disasterous.Longitude,
                                    lat = Disasterous.Latitude,
                                    text = Disasterous['Codice_Unico'],
                                    marker_color='brown',
                                    name='4',
                                    marker_size=8,
                                                  ),
            ],
            'layout': go.Layout(
                title = 'Best Western Hotels',
                legend={"title":"Earthquake Zones"},
                hovermode='closest',
                plot_bgcolor='black',
                height=1000,
                width=1000,
                geo = dict(
                    scope='europe',
                    resolution = 50,
                    countrycolor = 'blue',
                    landcolor="#DDDDDD",
                    showland=True,
                 )
            )
        })],
    style={'width':'50%','height':'100%', 'float':'left'}),
         html.Div([  # this Div contains our output graph and vehicle stats
    dcc.Graph(
        id='Danni_Diretti',
        figure={
            'data': [dict(name='Valore', x=['Fabbricato', 'Contenuto'], y=[0,0],type='bar'),
            dict(name='Terremoto', x=['Fabbricato', 'Contenuto'], y=[0,0],type='bar'),
            dict(name='Inondazione, Alluvione, Allagamento', x=['Fabbricato', 'Contenuto'],y=[0,0],type='bar'),
            dict(name='Terrorismo', x=['Fabbricato', 'Contenuto'], y=[0,0],type='bar'),
            dict(name='Eventi Socio Politici, Atti Dolosi', x=['Fabbricato', 'Contenuto'], y=[0,0],type='bar')],
            'layout': go.Layout(
                title = 'Danni Diretti',
                barmode='group',
                 bargroupgap=0.1,
                 height=1000,
            )
        }
    ),

    ],style={'width':'50%', 'float':'right'})])

@app.callback(
    Output('Danni_Diretti', 'figure'),
    [Input('mpg_scatter', 'hoverData')])
def callback_graph(hoverData):
    ii =  hoverData['points'][0]['text']
    i=df[df['Codice_Unico']==ii].index[0]
    fig = {
        'data': [dict(name='Valore', x=['Fabbricato', 'Contenuto'], y=MP.loc[i],type='bar'),
        dict(name='Terremoto', x=['Fabbricato', 'Contenuto'], y=MP.loc[i]*df.loc[i]['Coef. Terremoto'],type='bar'),
        dict(name='Inondazione, Alluvione, Allagamento', x=['Fabbricato', 'Contenuto'], y=MP.loc[i]*df.loc[i]['Coef.  Inondazione, Alluvione, Allagamento'],type='bar'),
        dict(name='Terrorismo', x=['Fabbricato', 'Contenuto'], y=MP.loc[i]*df.loc[i]['Coef. Terrorismo'],type='bar'),
        dict(name='Eventi Socio Politici, Atti Dolosi', x=['Fabbricato', 'Contenuto'], y=MP.loc[i]*df.loc[i]['Coef. Eventi Socio Politici, Atti Dolosi'],type='bar'),
        ],
        'layout': go.Layout(
            title = ii,
            height = 1000,
            width=1000,
        )
    }
    return fig



app.run_server(debug=False, use_reloader=False)
