from dash import *
from dash.exceptions import PreventUpdate
import dash_leaflet as dl
import os
import psycopg2 as pg
import geopandas as gpd
import pandas as pd
from branca.element import MacroElement
from jinja2 import Template
import json


connection = pg.connect(
    database=DATABASE,
    user=USER,
    password=PASSWORD,
    host=HOST,
    port=PORT
)

connection.autocommit = False
cursor = connection.cursor()
query = "SELECT name_en, geom from polygon"
df = gpd.read_postgis(sql=query, con=connection)


data = pd.read_excel(r'D:\\ИЦТО\\map\detsad.xlsx')
data_column_list = data.columns


parse_json_list = []

for _, r in df.iterrows():
    sim_geo = gpd.GeoSeries(r['geom']).simplify(tolerance=0.001)
    parse_json = sim_geo.to_json()
    parse_json = json.loads(parse_json)
    parse_json_list.append(parse_json)


polygon_list = []
for _, r in df.iterrows():
    sim_geo = gpd.GeoSeries(r['geom']).simplify(tolerance=0.001)
    parse_json = sim_geo.to_json()
    parse_json = json.loads(parse_json)
    marker_list = []

    for i in range(len(data)):              #Вывод всех маркеров
    # for i in range(0,2):                      #Первые 2 маркера - Волоколамский городской округ
        excel = data[data_column_list[0]][i].lower()
        name = r['name_en'].lower()
        if name.find(excel) != -1:
            marker_list.append(dl.Marker(
                position=[data[data_column_list[5]][i], data[data_column_list[6]][i]], 
                children=dl.Tooltip(data[data_column_list[1]][i]), 
                id={'type': 'marker', 'index': str(i)}, n_clicks=0
                ))
        else:
            continue

    polygon_list.append(
        dl.Overlay(dl.LayerGroup(children=[
            dl.GeoJSON(data=parse_json, zoomToBounds=True, zoomToBoundsOnClick=True, children=[
                marker_list[i] for i in range(len(marker_list))
            ])
        ]), name=r['name_en'], checked=False)
    )


app = dash.Dash(suppress_callback_exceptions=True)
app.layout = html.Div([
    dl.Map(children=[
        dl.LayersControl(
            [dl.BaseLayer(dl.TileLayer(), checked=True)] +
            [polygon_list[i] for i in range(len(polygon_list))], 
        )
    ], center=[55.7522, 37.6156], zoom=5),
    html.Div([
        html.Div(id='info', children=[]),
    ])
], id='map', style={'width': '99vw', 'height': '90vh'})\

@app.callback(
    Output('info', 'children'),
    Input({'type': 'marker', 'index': ALL}, 'n_clicks'),
    State('info', 'children'),
)
def display_dropdowns(n_clicks, children):
    trigger = ctx.triggered_id
    # print(trigger, trigger['index'])
    if not trigger:
        return no_update
    else:
        # print(id)
        # print(n_clicks)
        new_element = html.Div(
            id={
                'type': 'dynamic-output',
                # 'index': id['index']
                'index': trigger['index']
            })
        if len(children) > 0:
            children.pop(0)
        children.append(new_element)
        return children


@app.callback(
    Output({'type': 'dynamic-output', 'index': MATCH}, 'children'),
    [Input({'type': 'marker', 'index': MATCH}, "position"),
    Input({'type': 'marker', 'index': MATCH}, "n_clicks")],
    State({'type': 'marker', 'index': MATCH}, 'id'),
)
def display_output(position, n_clicks, id):
    if n_clicks == 0:
        raise PreventUpdate
    else:
        # print(n_clicks)
        return html.Div('Dropdown {} = {} and {}'.format(id['index'], position, n_clicks))


if __name__ == '__main__':
    app.run_server(debug=True)
