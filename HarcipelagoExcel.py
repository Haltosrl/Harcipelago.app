import dash
from dash import dcc, html, Input, Output
import pandas as pd
import openpyxl
import plotly.graph_objs as go
import datetime

# Caricamento del file Excel
excel_file = '/Users/gp/Library/CloudStorage/OneDrive-Personale/Meravimutuo/Harcipelago/Harcipelago Excel2.xlsx'
df = pd.read_excel(excel_file, sheet_name='Tabella')
df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
df_totali = pd.read_excel(excel_file, sheet_name='Totali', header=None)
df_harcipelago = pd.read_excel(excel_file, sheet_name='Harcipelago', usecols='B:I', skiprows=1, nrows=35)
df_harcipelago = df_harcipelago[df_harcipelago['Anni'].between(2010, 2023)]
print(df_totali.head())  # Debug: stampa i primi 5 valori del foglio Totali per controllo

# Filtrare i dati dal foglio Tabella per ottenere l'intervallo di 10 anni
def filter_10_years_data(start_date):
    end_date = start_date + pd.DateOffset(years=10)
    filtered_df = df[(df['Data'] >= start_date) & (df['Data'] <= end_date)]
    return filtered_df

# Inizializzazione dell'app Dash
app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "Simulazione Harcipelago"

app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Header([
        html.Nav([
            dcc.Link('Home', href='/', className='nav-link', style={'marginRight': '20px', 'color': '#d3d3d3', 'textDecoration': 'none'}),
            dcc.Link('Harcipelago', href='/harcipelago', className='nav-link', style={'marginRight': '20px', 'color': '#d3d3d3', 'textDecoration': 'none'}), dcc.Link('Investimento', href='/simulazione', className='nav-link', style={'marginRight': '20px', 'color': '#d3d3d3', 'textDecoration': 'none'})
        ], style={'textAlign': 'right', 'padding': '20px', 'backgroundColor': '#333333', 'fontSize': '26px', 'borderRadius': '10px', 'color': '#d3d3d3'})
    ]),
    html.Div(id='page-content', style={'padding': '40px', 'maxWidth': '1200px', 'margin': '0 auto'})
], style={'fontFamily': 'Verdana, sans-serif', 'backgroundColor': '#2b2b2b', 'padding': '40px', 'color': '#d3d3d3'})


@app.callback(
    Output('page-content', 'children'),
    Input('url', 'pathname')
)
def display_page(pathname):
    if pathname == '/simulazione':
        return html.Div([
            html.H1('Investimento Harcipelago per 10 Anni', style={'textAlign': 'center', 'color': '#1abc9c', 'marginBottom': '30px', 'fontSize': '36px', }),
            html.Label('Seleziona la data di inizio:', style={'fontSize': '20px', 'fontWeight': 'bold', 'display': 'block', 'marginBottom': '10px'}),
            dcc.Dropdown(
                id='start-date-dropdown',
                options=[{'label': str(year), 'value': str(year)} for year in range(1990, 2015)],
                value='2014',
                style={'marginBottom': '20px', 'width': '200px', 'color': '#006400'}
            ),
            html.Div(id='summary-output', style={'fontSize': '20px', 'marginBottom': '20px', 'backgroundColor': '#3b3b3b', 'padding': '20px', 'borderRadius': '10px', 'boxShadow': '0 4px 8px rgba(0, 0, 0, 0.3)', 'color': '#e0e0e0'}),
            dcc.Graph(id='investment-graph', config={'displayModeBar': False}, style={'marginBottom': '30px', 'boxShadow': '0 4px 8px rgba(0, 0, 0, 0.1)'}),
            html.H3('Guadagni Annuali', style={'textAlign': 'center', 'color': '#d3d3d3', 'marginTop': '50px', 'fontSize': '30px', 'fontWeight': 'bold'}),
            html.Table(id='yearly-gain-table', style={'width': '100%', 'margin': '0 auto', 'borderCollapse': 'collapse', 'backgroundColor': '#3b3b3b', 'border': '1px solid #444', 'boxShadow': '0 4px 8px rgba(0, 0, 0, 0.3)', 'marginTop': '20px', 'padding': '10px', 'color': '#e0e0e0'})
        ])
    elif pathname == '/harcipelago':
        return html.Div([
            dcc.Graph(
                id='gain-histogram',
                figure={
                    'data': [
                        go.Bar(
                            x=df_harcipelago[df_harcipelago['Anni'].between(1999, 2023)]['Anni'],
                            y=df_harcipelago[df_harcipelago['Anni'].between(1999, 2023)]['Gain'],
                            marker=dict(color='#1abc9c')
                        )
                    ],
                    'layout': {
                        'title': 'Anni vs Gain', 'titlefont': {'color': '#d3d3d3'},
                        'xaxis': {'title': 'Anni', 'titlefont': {'color': '#d3d3d3'}, 'tickfont': {'color': '#d3d3d3'}},
                        'yaxis': {'title': 'Gain', 'titlefont': {'color': '#d3d3d3'}, 'tickfont': {'color': '#d3d3d3'}},
                        'plot_bgcolor': '#2b2b2b',
                        'paper_bgcolor': '#333333'
                    }
                },
                style={'marginBottom': '30px', 'boxShadow': '0 4px 8px rgba(0, 0, 0, 0.1)'}
            ),
            html.H1('Harcipelago', style={'textAlign': 'center', 'color': '#1abc9c', 'marginBottom': '20px', 'fontSize': '36px'}),
            html.P("Tabella annuale Harcipelago.", style={'fontSize': '20px', 'textAlign': 'center', 'maxWidth': '800px', 'margin': '0 auto', 'lineHeight': '1.6'}),
            html.Table(
                [html.Thead(html.Tr([html.Th(col, style={'border': '1px solid #444', 'padding': '8px', 'backgroundColor': '#3b3b3b', 'color': '#1abc9c'}) for col in df_harcipelago.columns]))] +
                [html.Tr([html.Td(
                    html.A(str(df_harcipelago.iat[row, col]), href='/2023', style={'color': '#1abc9c', 'textDecoration': 'none'}) if df_harcipelago.columns[col] == 'Anni' and df_harcipelago.iat[row, col] == 2023 else str(df_harcipelago.iat[row, col]) if df_harcipelago.columns[col] == 'Anni' else f"{df_harcipelago.iat[row, col]:,.2f}",
                    style={'border': '1px solid #444', 'padding': '8px', 'color': 'red' if df_harcipelago.iat[row, col] < 0 and df_harcipelago.columns[col] != 'Anno' else '#e0e0e0', 'textAlign': 'right', 'backgroundColor': '#4b4b4b' if df_harcipelago.columns[col] == 'Gain' else '#3b3b3b'}) for col in range(df_harcipelago.shape[1])]) for row in range(df_harcipelago.shape[0])],
                style={'width': '100%', 'margin': '0 auto', 'borderCollapse': 'collapse', 'backgroundColor': '#3b3b3b', 'border': '1px solid #444', 'boxShadow': '0 4px 8px rgba(0, 0, 0, 0.3)', 'marginTop': '20px', 'padding': '10px', 'color': '#e0e0e0'}
            )
        ])
    elif pathname == '/2023':
        return html.Div([
            html.H1("Dettagli per l'anno 2023", style={'textAlign': 'center', 'color': '#1abc9c', 'marginBottom': '20px', 'fontSize': '36px'}),
            html.P("Dettagli relativi all'anno 2023 saranno qui visualizzati.", style={'fontSize': '20px', 'textAlign': 'center', 'maxWidth': '800px', 'margin': '0 auto', 'lineHeight': '1.6', 'color': '#e0e0e0'})
        ])
    else:
        return html.Div([
            html.H1('Benvenuti nella Simulazione Harcipelago', style={
                'textAlign': 'center', 'color': '#1abc9c', 'marginBottom': '30px', 'fontSize': '36px'
            }),
            html.P('Questo sito permette di simulare l\'investimento Harcipelago e visualizzare i dati disponibili.', style={
                'fontSize': '20px', 'textAlign': 'center', 'maxWidth': '800px', 'margin': '0 auto', 'lineHeight': '1.6'
            }),
            dcc.Graph(
                id='home-line-graph',
                figure={
                    'data': [
                        go.Scatter(
                            x=df['Data'],
                            y=df['Close'],
                            mode='lines',
                            line=dict(color='#1abc9c')
                        )
                    ],
                    'layout': {
                        'title': {'text': 'Indice', 'font': {'color': '#d3d3d3'}},
                        'xaxis': {'title': 'Data', 'titlefont': {'color': '#d3d3d3'}, 'tickfont': {'color': '#d3d3d3'}},
                        'yaxis': {'title': 'Close', 'titlefont': {'color': '#d3d3d3'}, 'tickfont': {'color': '#d3d3d3'}},
                        'plot_bgcolor': '#2b2b2b',
                        'paper_bgcolor': '#333333'
                    }
                },
                style={'marginBottom': '30px', 'boxShadow': '0 4px 8px rgba(0, 0, 0, 0.1)'}
            )
        ])


@app.callback(
    [Output('investment-graph', 'figure'),
     Output('summary-output', 'children'),
     Output('yearly-gain-table', 'children')],
    [Input('start-date-dropdown', 'value')]
)
def update_simulation(start_year):
    if start_year is None:
        return dash.no_update

    start_date = pd.to_datetime(f'{start_year}-01-01')
    filtered_df = filter_10_years_data(start_date)

    # Calcolo del Guadagno/Perdita Finale e Capitale Investito
    total_gain = filtered_df['Daily Gain'].sum()
    capital_invested = filtered_df['Capitale da Reinvestire'].abs().sum() + df_totali.iat[2, 2] + (df_totali.iat[3, 2] * 12)

    investment_summary = [
        html.P(f"Capitale Investito Totale: {capital_invested:,.2f} €"),
        html.P(f"Guadagno/Perdita Finale: {total_gain:,.2f} €", style={'color': 'red' if total_gain < 0 else '#d3d3d3'})
    ]

    # Grafico
    figure = {
        'data': [
            go.Scatter(x=filtered_df['Data'], y=filtered_df['Daily Gain'].cumsum(), mode='lines', line=dict(color='#1abc9c'))
        ],
        'layout': {
            'title': {'text': "Andamento dell'Investimento", 'font': {'color': '#d3d3d3'}},
            'xaxis': {'title': 'Data', 'titlefont': {'color': '#d3d3d3'}, 'tickfont': {'color': '#d3d3d3'}},
            'yaxis': {'title': 'Guadagno Cumulato', 'titlefont': {'color': '#d3d3d3'}, 'tickfont': {'color': '#d3d3d3'}},
            'plot_bgcolor': '#2b2b2b',
            'paper_bgcolor': '#333333'
        }
    }

     # Tabella Guadagni Annuali
    yearly_gain = filtered_df.resample('Y', on='Data').sum()
    yearly_gain['Cumulative Gain'] = yearly_gain['Daily Gain'].cumsum()
    table_header = [html.Thead(html.Tr([html.Th('Anno', style={'border': '1px solid #444', 'padding': '8px', 'backgroundColor': '#3b3b3b', 'color': '#1abc9c'}), html.Th('Guadagno Annuale', style={'border': '1px solid #444', 'padding': '8px', 'backgroundColor': '#3b3b3b', 'color': '#1abc9c'}), html.Th('Guadagno Cumulato', style={'border': '1px solid #444', 'padding': '8px', 'backgroundColor': '#3b3b3b', 'color': '#1abc9c'})]))]
    table_body = [html.Tr([
        html.Td(year.year, style={'border': '1px solid #444', 'padding': '8px', 'textAlign': 'right'}),
        html.Td(f"{row['Daily Gain']:,.2f}", style={'border': '1px solid #444', 'padding': '8px', 'textAlign': 'right', 'color': 'red' if row['Daily Gain'] < 0 else '#e0e0e0'}),
        html.Td(f"{row['Cumulative Gain']:,.2f}", style={'border': '1px solid #444', 'padding': '8px', 'textAlign': 'right', 'color': 'red' if row['Cumulative Gain'] < 0 else '#e0e0e0'})
    ]) for year, row in yearly_gain.iterrows()]

    return figure, investment_summary, table_header + table_body

if __name__ == '__main__':
    app.run_server(debug=True)
