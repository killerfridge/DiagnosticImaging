import dash
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output
import pandas as pd
from philips_colors import tableau_hex as tableau
import plotly.graph_objs as go


def table_view(df: pd.DataFrame, max_rows=10)->html.Table:
    # Header
    return html.Table([html.Tr([html.Th(col) for col in df.columns])] + [html.Tr([
        html.Td(df.iloc[i][col]) for col in df.columns
        ]) for i in range(min(len(df), max_rows))])


df_di = pd.read_excel('data\\df_waiting_list.xlsx')

df_di = df_di.groupby(['STP', 'Period']).sum().reset_index()

df_di_table = df_di[[
    'STP',
    'Period',
    "('Total Activity', 'CT')",
    "('Total Activity', 'MRI')",
    "('Total WL', 'CT')",
    "('Total WL', 'MRI')",
]].copy()

df_di_table.columns = [
    'STP',
    'Period',
    'CT Activity',
    'MRI Activity',
    'CT Waiting List',
    'MRI Waiting List',
]

stylesheets = ['https://codepen.io/chriddyp/pen/bWLwgP.css']

app = dash.Dash(external_stylesheets=stylesheets)

options = [
    {'label': x, 'value': x} for x in df_di_table['STP'].unique()
]


app.layout = html.Div(
    style={
        'fontFamily': 'Helvetica'
    },
    children=[
        html.H1('NHS STP CT and MRI Activity and Waiting Lists', style={'textAlign': 'center'}),
        html.Div(className='row', children=[
            dcc.Dropdown(
                id='stp_selection',
                options=options,
                multi=False,
                value='Bath, Swindon and Wiltshire',
                className='five columns',
            ),
            dcc.Input(id='rolling_window', value=3, className='two columns'),
            ]),
        html.Div(className='row', children=[
            dcc.Graph(className='six columns', id='CT_chart'),
            dcc.Graph(className='six columns', id='MR_chart')
        ]),
        html.Div(className='row', children=[
            dcc.Graph(className='six columns', id='CT_overall'),
            dcc.Graph(className='six columns', id='MR_overall'),
        ])
    ]
)


@app.callback(
    Output('CT_chart', 'figure'),
    [Input('stp_selection', 'value'),
     Input('rolling_window', 'value')]
)
def update_ct(stp, window):
    try:
        window = int(window)
    except ValueError:
        window = 1
    df_filtered = df_di_table[df_di_table['STP'] == stp].groupby('Period').sum().rolling(
        window=window,
        min_periods=window,
    ).mean().reset_index().dropna()

    return {
        'data': [
            {'x': df_filtered.loc[:, 'Period'], 'y': df_filtered['CT Activity'], 'name': 'CT Activity',
             'type': 'bar'},
            {'x': df_filtered.loc[:, 'Period'], 'y': df_filtered['CT Waiting List'], 'name': 'CT Waiting List',
             'type': 'bar'},
        ],
        'layout': {
            'title': f'CT Activity and Waiting List<br>{stp}'
        }
    }


@app.callback(
    Output('MR_chart', 'figure'),
    [Input('stp_selection', 'value'),
     Input('rolling_window', 'value')]
)
def update_mr(stp, window):
    try:
        window = int(window)
    except ValueError:
        window = 1

    df_filtered = df_di_table[df_di_table['STP'] == stp].groupby('Period').sum().rolling(
        window=window
    ).mean().reset_index().dropna()

    return {
        'data': [
            {'x': df_filtered.loc[:, 'Period'], 'y': df_filtered['MRI Activity'], 'name': 'MRI Activity',
             'type': 'bar'},
            {'x': df_filtered.loc[:, 'Period'], 'y': df_filtered['MRI Waiting List'], 'name': 'MRI Waiting List',
             'type': 'bar'},
        ],
        'layout': {
            'title': f'MRI Activity and Waiting List<br>{stp}'
        }
    }


@app.callback(
    Output('CT_overall', 'figure'),
    [Input('stp_selection', 'value'),
     Input('rolling_window', 'value')]
)
def overall_ct(stp, window):

    try:
        window = int(window)
    except ValueError:
        window = 1

    overall_df = df_di_table.groupby(['STP', 'Period']).sum()
    overall_df = overall_df.groupby(level='STP').apply(lambda x: x.rolling(window).mean()).dropna()
    overall_df = overall_df.groupby(level='STP').last().sort_values('CT Activity', ascending=False)

    colors = [tableau[2] if x == stp else tableau[0] for x in overall_df.index]

    return {
        'data': [
            {'x': overall_df.index, 'y': overall_df['CT Activity'], 'type': 'bar', 'marker': {'color': colors},
             'name': 'Activity'},
            {
                'x': overall_df.index,
                'y': [overall_df['CT Activity'].mean() for _ in overall_df.index],
                'type': 'scatter',
                'line': {
                    'dash': 'dot',
                    'color': 'grey'
                },
                'mode': 'lines',
                'name': 'Average Activity'
            }
        ],
        'layout': {
            'title': 'CT Activity'
        }
    }


@app.callback(
    Output('MR_overall', 'figure'),
    [Input('stp_selection', 'value'),
     Input('rolling_window', 'value')]
)
def overall_mr(stp, window):

    try:
        window = int(window)
    except ValueError:
        window = 1

    overall_df = df_di_table.groupby(['STP', 'Period']).sum()
    overall_df = overall_df.groupby(level='STP').apply(lambda x: x.rolling(window).mean()).dropna()
    overall_df = overall_df.groupby(level='STP').last().sort_values('MRI Activity', ascending=False)

    colors = [tableau[2] if x == stp else tableau[0] for x in overall_df.index]

    return {
        'data': [
            {'x': overall_df.index, 'y': overall_df['MRI Activity'], 'type': 'bar', 'marker': {'color': colors},
             'name': 'Activity'},
            go.Scatter(
                x=overall_df.index,
                y=[overall_df['MRI Activity'].mean() for _ in overall_df.index],
                name='Average Activity',
                line={
                    'dash': 'dot',
                    'color': 'grey',
                },
                mode='lines',
            ),
        ],
        'layout': {
            'title': 'MRI Activity'
        }
    }


if __name__ == '__main__':
    app.run_server(debug=True)

