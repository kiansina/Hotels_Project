from dash import dcc, html, Input, Output
import dash
import plotly.graph_objs as go
import pandas as pd

external_stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']
df=pd.read_excel(r"C:\Users\sina.kian\Desktop\Ricardo\updates\DataBase_5_final.xlsx")

cc=['green', 'yellow', 'orange', 'red', 'brown']
S=['Denominazione Hotel','Indirizzo (Ubicazione Hotel)','Provincia \n(Ubicazione Hotel)','Fabbricato',
       'Contenuto','Terremoto, Inondazione, Alluvione, Allagamento','Coef. Terremoto', 'Earthquake Zone',  'Earthquake_Risk_Score_Risk_Score', 'Flood_Risk_Score_Risk_Score','Longitude','Latitude','Codice Hotel','Codice_Unico']
#S=['Codice Hotel','Indirizzo (Ubicazione Hotel)',
#        'Provincia \n(Ubicazione Hotel)', 'Città (Ubicazione Hotel)',
#        'Cap (Ubicazione Hotel)','Terremoto, Inondazione, Alluvione, Allagamento','Coef. Terremoto','Earthquake Zone', 'Earthquake_Risk_Score_Risk_Score', 'Flood_Risk_Score_Risk_Score','Longitude','Latitude']
colors = {
    'background': '#111111',
    'text': '#7FDBFF'
}
dfs=df
cc=['green', 'yellow', 'orange', 'red', 'brown']
#dfs['point_c']=['green']*len(df)


#for i in dfs.index:
#    if dfs['Earthquake Zone'][i]==0:
#        dfs['point_c'][i]=cc[0]
#    if dfs['Earthquake Zone'][i]==1:
#        dfs['point_c'][i]=cc[1]
#    if dfs['Earthquake Zone'][i]==2:
#        dfs['point_c'][i]=cc[2]
#    if df['Earthquake Zone'][i]==3:
#        dfs['point_c'][i]=cc[3]
#    if dfs['Earthquake Zone'][i]==4:
#        dfs['point_c'][i]=cc[4]






MP=df[['Fabbricato', 'Contenuto']]
dfs=df[S]

ez='0'

app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

app.layout = html.Div([
    html.Div(
               html.H1('Hello Dear Best Western Group',
               style={
                       'textAlign': 'center',
                       'color': colors['text']
                   })),
    html.Div(
           html.H4('Please select the Earth quake zone in which the hotels are placed',
           style={
               'textAlign': 'center',
               'color': colors['text']
           }
       )),
    html.Div(children=[
                  html.H4('',
                        style={
                            'textAlign': 'center',
                            'color': colors['text']
                        }
                        ),
    dcc.RadioItems(['0', '1','2', '3', '4'], '0', id='my-slider', inline=True,  style={'color': colors['text']})], style={'display':'flex','Align': 'center', 'justifyContent':'center'}),
    html.Data(id='slider-output-container', value=ez),

    html.Div([  # this Div contains our scatter plot
    dcc.Graph(
        id='mpg_scatter',
        figure={
            'data': [go.Scattermapbox(
                                mode='markers',
                                lat = dfs[dfs['Earthquake Zone']==int(ez)].Latitude,
                                lon = dfs[dfs['Earthquake Zone']==int(ez)].Longitude,
                                text=dfs[dfs['Earthquake Zone']==int(ez)]['Codice_Unico'],
                                marker=dict(size=10,
                                color=cc[int(ez)])
            )],
            'layout': go.Layout(
            mapbox = {
                'style': "carto-positron",
                'center': { 'lon': 12.5, 'lat': 41.9},
                'zoom': 5.5,
            },
                title = 'Hotel Locations based on Earthquake Zones',
                margin={"r":0,"t":0,"l":0,"b":0},
                hovermode='closest',
                showlegend = False,
                height=1000,
                width=1000,
                plot_bgcolor= colors['background'],
                paper_bgcolor= colors['background'],
                font= {
                    'color': colors['text']
                    },
            )
        })],
    style={'width':'50%','height':'100%', 'float':'left', 'backgroundColor': colors['background']}),


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
                 plot_bgcolor= colors['background'],
                 paper_bgcolor= colors['background'],
                 font= {
                     'color': colors['text']
                     },
            )
        }
    ),

    ],style={'width':'50%', 'float':'right'})],style={'backgroundColor': colors['background']})







@app.callback(
    Output('slider-output-container', 'value'),
    Input('my-slider', 'value'))
def update_output(value):
    ez=str(value)
    return ez


@app.callback(
    Output('mpg_scatter', 'figure'),
    [Input('slider-output-container', 'value')])
def update_output(value):
    Ss=int(value)
    fig={
        'data': [go.Scattermapbox(
                            mode='markers',
                            lat = dfs[dfs['Earthquake Zone']==Ss].Latitude,
                            lon = dfs[dfs['Earthquake Zone']==Ss].Longitude,
                            text=dfs[dfs['Earthquake Zone']==Ss]['Codice_Unico'],
                            marker=dict(size=10,
                            color=cc[Ss])
        )],
        'layout': go.Layout(
        mapbox = {
            'style': "carto-positron",
            'center': { 'lon': 12.5, 'lat': 41.9},
            'zoom': 5.5,
        },
            title = 'Hotel Locations based on Earthquake Zones',
            margin={"r":0,"t":0,"l":0,"b":0},
            hovermode='closest',
            showlegend = False,
            height=1000,
            width=1000,
            plot_bgcolor= colors['background'],
            paper_bgcolor= colors['background'],
            font= {
                'color': colors['text']
                },
        )
    }
    return fig

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
            plot_bgcolor= colors['background'],
            paper_bgcolor= colors['background'],
            font= {
                'color': colors['text']
                },
        )
    }
    return fig





app.run_server(debug=True, use_reloader=False)
