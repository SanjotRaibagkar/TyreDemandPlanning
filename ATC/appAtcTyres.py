# standard library
import os

# dash libs
import dash
from dash.dependencies import Input, Output, State, Event
import dash_core_components as dcc
import dash_html_components as html
import plotly.figure_factory as ff
import plotly.graph_objs as go

# pydata stack
import pandas as pd
from sqlalchemy import create_engine

from datetime import datetime as dt

import io
import flask
from flask import send_file, url_for, request
import urllib.parse as p
import xlsxwriter
import zipfile
import io
import pathlib
import numpy as np

# set params
#conn = create_engine(os.environ['DB_URI'])
conn = create_engine("sqlite:///atcTyres.db")
globalStockReport = pd.DataFrame(columns = ['plantCode', 'rmCode', 'calDate', 'stock', 'qty', 'etsDate'])

#######################
# Data Analysis / Model
#######################

def fetch_data(q):
    print ("Executing Query:", q)	
    result = pd.read_sql(
        sql=q,
        con=conn
    )
    return result


def get_plants():
    '''Returns the list of plants that are stored in the database'''

    plant_query = (
        '''
        SELECT DISTINCT plantCode
        FROM StockReport
        '''
    )
    plants = fetch_data(plant_query)
    plants = list(plants['plantCode'].sort_values(ascending=True))
    return plants


def get_rawMats(plant):
    '''Returns the raw material of selected plant'''

    rawMats_query = (
        '''
        SELECT DISTINCT rmCode
        FROM StockReport
        WHERE plantCode='%s'
        '''%plant
    )
    rawMaterials = fetch_data(rawMats_query)
    rawMaterials = list(rawMaterials['rmCode'].sort_values(ascending=False))
    return rawMaterials 

def get_SafeStock(plantCode,rmCode):
    '''Returns the Safe Stock for selected RM and Plant'''

    safeStock_query = (
        '''
        SELECT safeStock
        FROM SafeStock
        WHERE plantCode='%s'
        AND rmCode='%s'
        '''%(plantCode,rmCode)
    )
    safeStockQty = fetch_data(safeStock_query)
    return safeStockQty['safeStock']

#DB query to fetch consumption history 
def get_ConsumptionHistory(plantCode, rmCode, reportDate):
    '''Returns consumption history for specified plant, rawMat since specified date'''

    report_query = (
        '''
        SELECT plantCode, rmCode, monthYear, mthlyConsumption
        FROM MthlyConsumption
        WHERE plantCode='%s'
        AND rmCode='%s'
        ORDER BY monthYear ASC
        '''%(plantCode, rmCode)
    )
    consumptionReport = fetch_data(report_query)
    return consumptionReport
        #AND monthYear >='{reportDate}'

#DB query to fetch stock report
def get_StockReport(plantCode, rmCode, reportDate):
    '''Returns stock details for specified plant, rawMat since specified date'''

    report_query = (
        '''
        SELECT plantCode, rmCode, calDate, stock, qty, etsDate 
        FROM StockReport
        WHERE plantCode='%s'
        AND rmCode='%s'
        AND calDate >='%s'
        ORDER BY calDate ASC
        '''%(plantCode, rmCode, reportDate)
    )
    stockReport = fetch_data(report_query)
    stockReport =modify_stock_report(stockReport,0,plantCode, rmCode, reportDate)
    return stockReport

def modify_stock_report (results , safeStock,plantCode, rmCode, reportDate):
    # temp hardcode  consumption number
    consumptionReport = get_ConsumptionHistory(plantCode, rmCode, reportDate)
    mnonthly_cosumption = consumptionReport['mthlyConsumption'].iloc[0]
    daily_cosumption = -mnonthly_cosumption/30
    stock = results['stock'].iloc[0]
    start_value = stock
    stop_value = stock +120*daily_cosumption
    #print("series vale" , start_value,stop_value,daily_cosumption)
   
    data = np.arange(start_value,stop_value,daily_cosumption)
    stock_series = pd.Series(data)
    #print("Stock series is ",len(stock_series))
   
    date  = results['calDate'].iloc[0]
   
    plantcode = results['plantCode'].iloc[0]
    rmcode  =results['rmCode'].iloc[0]
    dataframe_new = pd.DataFrame(index=range(120),columns = ['calDate','plantCode', 'rmCode', 'stock', 'qty', 'etsDate'])
    dataframe_new['plantCode'] = plantcode
    dataframe_new['rmCode'] = rmcode
    dataframe_new['stock'].iloc[0] = stock
    dataframe_new['qty'] = pd.to_numeric(results['qty'])
    dataframe_new['etsDate'] = results['etsDate']
   
    dataframe_new['qty'].fillna(0, inplace=True)
    #print("Report date " ,  pd.to_numeric(dataframe_new['qty']))

    rng = pd.date_range(date, periods=120, freq='D')

   # print("data range datafram", rng.to_series)
    dataframe_new =  dataframe_new.assign(calDate= rng)

    dataframe_new =  dataframe_new.assign(stock= stock_series)

   # print("concated frame", dataframe_new)
    dataframe_new['stock'] =dataframe_new['stock'] +pd.to_numeric(dataframe_new['qty'])
    dataframe_new['stock'] = dataframe_new['stock'].round(2)
    return dataframe_new

def draw_stockReport_graph(results , safeStock):
    dates = results['calDate']
    #safeStocks = [safeStock] * results.shape[0]
    safeStocks = pd.DataFrame([safeStock] * results.shape[0])
    stocks = results['stock']

    # Trace Safe Stock
    trace0 = go.Scatter(
        x = dates,
        y = safeStocks[0],
        name = 'Safe Stock',
        line = dict(
            color = ('rgb(205, 12, 24)'),
            width = 4)
    )
    # Trace Actual Stock
    trace1 = go.Scatter(
        x = dates,
        y = stocks,
        name = 'Actual Stock',
        line = dict(
            color = ('rgb(22, 96, 167)'),
            width = 4,)
    )

    figure = go.Figure(
        data= [ trace0 ,  trace1],
        layout=go.Layout(
            title='Stock Report',
            showlegend=True
        )
    )
    return figure


#########################
# Dashboard Layout / View
#########################

def generate_table(dataframe, max_rows=10):
    '''Given dataframe, return template generated using Dash components
    '''
    return html.Table(
        # Header
        [html.Tr([html.Th(col) for col in dataframe.columns])] +

        # Body
        [html.Tr([
            html.Td(dataframe.iloc[i][col]) for col in dataframe.columns
        ]) for i in range(min(len(dataframe), max_rows))]
    )


def onLoad_plant_options():
    '''Actions to perform upon initial page load'''

    plant_options = (
        [{'label': plantCode, 'value': plantCode}
         for plantCode in get_plants()]
    )
    return plant_options

# Set up Dashboard and create layout
app = dash.Dash(__name__)
app.css.append_css({
    "external_url": "https://codepen.io/chriddyp/pen/bWLwgP.css"
})

app.layout = html.Div([
    html.Div([
        html.Div([
        html.Img(src='/assets/ATGlogo.png',style={'width': '150px'}),
        ], className='three columns'),
        html.Div(className='nine columns'),
    ], className='twelve columns'),
    # Page Header
    html.Div([
        html.H4('Stock Report Viewer')
    ]),

    # Dropdown Grid
    html.Div([
        html.Div([
            # Select Plant Dropdown
            html.Div([
                html.Div('Select Plant', className='three columns'),
                html.Div(dcc.Dropdown(id='plant-selector',
                                      options=onLoad_plant_options()),
                         className='nine columns')
            ]),

            # Select Raw Material Dropdown
            html.Div([
                html.Div('Select Raw Material', className='three columns'),
                html.Div(dcc.Dropdown(id='rawMat-selector'),
                         className='nine columns')
            ]),
        ], className='six columns'),

        # Empty
        # html.Div(className='six columns'),

        html.Div([
             html.P('Select Date'),
             dcc.DatePickerSingle(
        id='reportStartDate',
        min_date_allowed=dt(2001, 8, 1),
        max_date_allowed=dt(2018, 10, 1),
        placeholder='Start Date',
        ),
        ], className='three columns'),

        #Link to trigger /download_excel route
        html.Div(children=[
            html.P(html.A(
                    'Download Data',
                    id='excel-download-link',
                    download="",
                    href="",
                    target="_blank")),
            html.P(html.A(
                    'Download Zip',
                    id='zip-download-link',
                    download="",
                    href="",
                    target="_blank",
		    n_clicks=0)),
            html.P(html.A(['Print PDF'],className="button no-print print")),
        ], className='three columns'),
    ], className='twleve columns'),
    html.Div([
            html.Div([
            # Safe Stock table
            html.Div(children='Safe Stock:',className='three columns'),
            html.Div(id='safe-stock',className='one columns'),
            html.Div(className='two columns'),
            ],className='six columns'),
            
	    html.Div([
            #Consumption History 
            html.Table(id='consumption-history'),
            ],className='six columns'),
    ],className='twelve columns'),

    # Stock Report Grid
    html.Div([

        # Stock Report Table
        html.Div(
            html.Table(id='stock-report'),
            className='six columns'
        ),

        # Stock Summary Table and Graph
        html.Div([
            # graph
            dcc.Graph(id='stock-graph')
            # style={},
        ], className='six columns')
    ]),
])

#############################################
# Interaction Between Components / Controller
#############################################

# Load Seasons in Dropdown
@app.callback(
    Output(component_id='rawMat-selector', component_property='options'),
    [
        Input(component_id='plant-selector', component_property='value')
    ]
)
def populate_rawMaterial_selector(plant):
    rawMaterials = get_rawMats(plant)
    return [
        {'label': rmCode, 'value': rmCode}
        for rmCode in rawMaterials
    ]


# Load Stock Report
@app.callback(
    Output(component_id='stock-report', component_property='children'),
    [
        Input(component_id='plant-selector', component_property='value'),
        Input(component_id='rawMat-selector', component_property='value'),
        Input(component_id='reportStartDate', component_property='date')
    ]
)
def load_stockReport(plantCode, rmCode, reportDate):
    results = get_StockReport(plantCode, rmCode, reportDate)
    return generate_table(results, max_rows=50)

# Load Safe Stock
@app.callback(
    Output(component_id='safe-stock', component_property='children'),
    [
        Input(component_id='plant-selector', component_property='value'),
        Input(component_id='rawMat-selector', component_property='value')
    ]
)
def load_stockReport(plantCode, rmCode):
    safeStock = get_SafeStock(plantCode, rmCode)
    return safeStock

# Load Consumption History
@app.callback(
    Output(component_id='consumption-history', component_property='children'),
    [
        Input(component_id='plant-selector', component_property='value'),
        Input(component_id='rawMat-selector', component_property='value'),
        Input(component_id='reportStartDate', component_property='date')
    ]
)
def load_consumptionHist(plantCode, rmCode, reportDate):
    consumptionHist = get_ConsumptionHistory(plantCode, rmCode, reportDate)
    return generate_table(consumptionHist, max_rows=50)

#Update Stock Point Graph
@app.callback(
    Output(component_id='stock-graph', component_property='figure'),
    [
        Input(component_id='plant-selector', component_property='value'),
        Input(component_id='rawMat-selector', component_property='value'),
        Input(component_id='reportStartDate', component_property='date')
    ]
)
def load_stockReport_graph(plantCode, rmCode, reportDate):
    results = get_StockReport(plantCode, rmCode, reportDate)
    safeStock = get_SafeStock(plantCode, rmCode)

    figure = []
    if len(results) > 0:
        figure = draw_stockReport_graph(results, safeStock)

    return figure

#Callback to download data as csv
#@app.callback(
#    Output(component_id='excel-download-link', component_property='href'),
#    [        Input(component_id='plant-selector', component_property='value'),
#        Input(component_id='rawMat-selector', component_property='value'),
#        Input(component_id='reportStartDate', component_property='date')]
#)
#def download_CSVSingle(plantCode, rmCode, reportDate):
#    stockReport = get_StockReport(plantCode, rmCode, reportDate)
#    globalStockReport = stockReport
#    globalStockReport= globalStockReport.dropna()    
#    csv_string = globalStockReport.to_csv(index=False, encoding='utf-8')
#    csv_string = "data:text/csv;charset=utf-8," +p.quote(csv_string)    
#    return csv_string

#Callback to download Excel for single RM
@app.callback(
    Output(component_id='excel-download-link', component_property='href'),
    [   Input(component_id='plant-selector', component_property='value'),
        Input(component_id='rawMat-selector', component_property='value'),
        Input(component_id='reportStartDate', component_property='date')]
)
def download_ExcelSingle(plantCode, rmCode, reportDate):
	return url_for('generate_singleRM_report',plantCode=plantCode,rmCode=rmCode,reportDate=reportDate)

#Callback to download ZIP file
@app.callback(
    Output('zip-download-link', 'href'),
    [Input('zip-download-link', 'n_clicks'),
    Input(component_id='plant-selector', component_property='value'),
    Input(component_id='reportStartDate', component_property='date')]
)
def generate_report_url(n_clicks, plantCode, reportDate):
	return url_for('generate_report_url',plantCode=plantCode,reportDate=reportDate)


# Flask routes to serve various files for download

# Serve Excel for single RM
@app.server.route('/dash/singleRMReportDownload')
def generate_singleRM_report():
	plantCode = request.args.get('plantCode') 
	rmCode = request.args.get('rmCode') 
	reportDate = request.args.get('reportDate') 
	stockReport = get_StockReport(plantCode, rmCode, reportDate)
	stockReport = stockReport.dropna()    
	rmDownload_filename = str(rmCode) + ".xlsx"
	wsName = str(plantCode) + "_" + str(rmCode)
	absolute_filename = os.path.join(os.getcwd(), rmDownload_filename)
	buf = io.BytesIO()
	#excel_writer = pd.ExcelWriter(absolute_filename, engine="xlsxwriter")
	excel_writer = pd.ExcelWriter(buf, engine="xlsxwriter")
	stockReport.to_excel(excel_writer, sheet_name=wsName)
	excel_writer.save()
	excel_data = buf.getvalue()
	buf.seek(0)
	#return send_file(absolute_filename, 
	return send_file(buf, 
		attachment_filename = rmDownload_filename, as_attachment = True)
		#mimetype = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

# Serve ZIP for All RM
@app.server.route('/dash/allRMZipdownload')
def generate_report_url():
        plantCode = request.args.get('plantCode') 
        reportDate = request.args.get('reportDate') 
        rawMaterials=get_rawMats(plantCode) 

        #Create  excelFile for each RM across the plant
        for rm in rawMaterials:
            rmReport = get_StockReport(plantCode,rm,reportDate)
            rmReport = rmReport.dropna()
            rm_filename = str(rm) + ".xlsx"
            wsName = str(plantCode) + "_" + str(rm)
            relative_filename = os.path.join('excelFiles',rm_filename)
            absolute_filename = os.path.join(os.getcwd(), relative_filename)
            writer = pd.ExcelWriter(absolute_filename, engine='xlsxwriter')
            rmReport.to_excel(writer, sheet_name = wsName)
            writer.save()
        zip_filename = str(plantCode) + ".zip"
        relative_filename = os.path.join('zipFiles',zip_filename)
        absolute_filename = os.path.join(os.getcwd(), relative_filename)
        excelFilePath = os.path.join(os.getcwd(), 'excelFiles')
        base_path = pathlib.Path('./excelFiles')
        with zipfile.ZipFile(relative_filename, mode='w') as z:
            for f_name in base_path.iterdir():
                z.write(f_name) 
         
        return send_file(relative_filename, attachment_filename = zip_filename, as_attachment = True)


external_css = ["https://cdnjs.cloudflare.com/ajax/libs/normalize/7.0.0/normalize.min.css",
                "https://cdnjs.cloudflare.com/ajax/libs/skeleton/2.0.4/skeleton.min.css",
                "//fonts.googleapis.com/css?family=Raleway:400,300,600",
                "https://codepen.io/bcd/pen/KQrXdb.css",
                "https://maxcdn.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css"]

for css in external_css:
    app.css.append_css({"external_url": css})

external_js = ["https://code.jquery.com/jquery-3.2.1.min.js",
               "https://codepen.io/bcd/pen/YaXojL.js"]

for js in external_js:
    app.scripts.append_script({"external_url": js})


# start Flask server
if __name__ == '__main__':
    app.run_server(
        debug=True,
        host='0.0.0.0',
        port=8050
    )
