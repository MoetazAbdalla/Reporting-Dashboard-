import dash
from dash import dcc, html
from dash.dependencies import Input, Output
import pandas as pd
import plotly.express as px

# Load the Excel file with one sheet
excel_file = 'Final-10.11.xlsx'
sheet_names = pd.ExcelFile(excel_file).sheet_names

# Include Bootstrap CSS as an external stylesheet for styling
external_stylesheets = ['https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css']

# Create a Dash app
app = dash.Dash(__name__, external_stylesheets=external_stylesheets)

# Dashboard Layout with Tabs
app.layout = html.Div([
    html.Div([
        html.H1('Istanbul MediPol University Reporting Dashboard',
                style={'textAlign': 'center', 'color': '#007BFF'}),

        # Tiles for metrics in a responsive grid
        html.Div([
            html.Div([  # Total Students
                html.H4('Total Students'),
                html.H2(id='total-students', className='metric-text'),
            ], className='col-lg-2 col-md-4 col-sm-6 mb-2 tile'),  # Reduce margin-bottom

            html.Div([  # Total Payments
                html.H4('Total Payments'),
                html.H2(id='total-payments', className='metric-text'),
            ], className='col-lg-2 col-md-4 col-sm-6 mb-2 tile'),  # Reduce margin-bottom

            html.Div([  # Top Nationality
                html.H4('Top Nationality'),
                html.H2(id='top-nationality', className='metric-text'),
            ], className='col-lg-2 col-md-4 col-sm-6 mb-2 tile'),  # Reduce margin-bottom

            html.Div([  # Top Program
                html.H4('Top Program'),
                html.H2(id='top-program', className='metric-text'),
            ], className='col-lg-2 col-md-4 col-sm-6 mb-2 tile'),  # Reduce margin-bottom

            html.Div([  # Top Status
                html.H4('Top Status'),
                html.H2(id='top-status', className='metric-text'),
            ], className='col-lg-2 col-md-4 col-sm-6 mb-2 tile'),  # Reduce margin-bottom

            html.Div([  # Performance
                html.H4('Performance'),
                html.H2(id='performance-metric', className='metric-text'),
            ], className='col-lg-2 col-md-4 col-sm-6 mb-2 tile'),  # Reduce margin-bottom
        ], className="row justify-content-center"),  # Center the row

        # Dropdown menu
        html.Div([
            html.Div([
                dcc.Dropdown(
                    id='agent-dropdown',
                    options=[{'label': sheet, 'value': sheet} for sheet in sheet_names],
                    value=sheet_names[0],  # Default value (first sheet)
                    className='custom-dropdown',  # Add a custom class for styling
                ),
            ], className='col-lg-12'),  # Ensures the dropdown spans the full width
        ], className='row menu-bar mb-4 container-fluid'),

        # Tabs for different visualizations
        dcc.Tabs([
            dcc.Tab(label='Overview', children=[
                html.Div([
                    # Row for pie charts (responsive flexbox)
                    html.Div([
                        html.Div([dcc.Graph(id='status-pie-chart', config={'displayModeBar': False})],
                                 className='col-lg-4 col-md-6 col-sm-12 mb-2 chart-container'),  # Reduce margin-bottom
                        html.Div([dcc.Graph(id='top-nationality-pie-chart', config={'displayModeBar': False})],
                                 className='col-lg-4 col-md-6 col-sm-12 mb-2 chart-container'),  # Reduce margin-bottom
                        html.Div([dcc.Graph(id='top-paid-countries-pie-chart', config={'displayModeBar': False})],
                                 className='col-lg-4 col-md-6 col-sm-12 mb-2 chart-container'),  # Reduce margin-bottom
                    ], className="row"),

                    # Row for top regions
                    html.Div([
                        html.Div([dcc.Graph(id='top-regions-bar-chart', config={'displayModeBar': False})],
                                 className='col-lg-6 col-md-12 mb-2 chart-container'),  # Reduce margin-bottom
                    ], className="row"),

                    # Pie charts for top 10 programs applied and total paid for top 10 programs
                    html.Div([
                        html.Div([dcc.Graph(id='top-programs-applied-pie-chart', config={'displayModeBar': False})],
                                 className='col-lg-6 col-md-12 mb-2 chart-container'),  # Reduce margin-bottom
                        html.Div([dcc.Graph(id='top-programs-paid-pie-chart', config={'displayModeBar': False})],
                                 className='col-lg-6 col-md-12 mb-2 chart-container'),  # Reduce margin-bottom
                    ], className="row"),

                    # New row for top 10 agents bar chart and regions performance pie chart
                    html.Div([
                        html.Div([
                            dcc.Graph(id='top-agents-bar-chart', config={'displayModeBar': False}),
                        ], className='col-lg-6 col-md-12 mb-2 chart-container'),  # Reduce margin-bottom

                        html.Div([  # Top paid regions pie chart
                            dcc.Graph(id='top-paid-regions-pie-chart', config={'displayModeBar': False}),
                        ], className='col-lg-6 col-md-12 mb-2 chart-container'),  # Reduce margin-bottom
                    ], className="row"),
                ], className="container-fluid")
            ]),
        ], className="tabs-container")
    ], className='main-container'),

    # Button at the bottom of the page for downloading as PDF
    html.Div([
        html.Button('Download as PDF', id='download-pdf', n_clicks=0, className='btn btn-primary download-btn'),
    ], className='fixed-bottom-btn'),

    # Hidden Div to include JavaScript for print functionality
    html.Div(id='pdf-js', style={'display': 'none'}),
])

# Callbacks to update the dashboard based on selected sheet (agent)
@app.callback(
    [Output('total-students', 'children'),
     Output('total-payments', 'children'),
     Output('top-nationality', 'children'),
     Output('top-program', 'children'),
     Output('top-status', 'children'),
     Output('performance-metric', 'children'),
     Output('status-pie-chart', 'figure'),
     Output('top-nationality-pie-chart', 'figure'),
     Output('top-paid-countries-pie-chart', 'figure'),
     Output('top-regions-bar-chart', 'figure'),
     Output('top-agents-bar-chart', 'figure'),
     Output('top-paid-regions-pie-chart', 'figure'),
     Output('top-programs-applied-pie-chart', 'figure'),
     Output('top-programs-paid-pie-chart', 'figure')],
    [Input('agent-dropdown', 'value')]
)
def update_dashboard(selected_sheet):
    # Load the data from the selected sheet
    df = pd.read_excel(excel_file, sheet_name=selected_sheet)

    # Check if the DataFrame is empty
    if df.empty:
        return "No data", "No data", "No data", "No data", "No data", "No data", {}, {}, {}, {}, {}, {}

    # Ensure the 'Date' column is properly formatted and in datetime format
    if 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], format='%d.%m.%Y', errors='coerce')

    # Top statistics
    total_students = df.shape[0]
    top_nationality = df['Nationality'].value_counts().idxmax() if 'Nationality' in df.columns else "N/A"
    top_program = df['Program'].value_counts().idxmax() if 'Program' in df.columns else "N/A"
    top_status = df['Status'].value_counts().idxmax() if 'Status' in df.columns else "N/A"
    total_paid = df['Status'].value_counts().get('Paid', 0)
    performance = (total_paid / total_students) * 100 if total_students > 0 else 0
    performance_metric = f"{performance:.2f}%"

    # Pie chart for status distribution
    status_counts = df['Status'].value_counts().reset_index()
    status_counts.columns = ['Status', 'Count']
    pie_chart = px.pie(status_counts, names='Status', values='Count', title="Status Distribution", hole=0.4)

    # Pie chart for top nationality distribution
    top_nationality_counts = df['Nationality'].value_counts().reset_index()
    top_nationality_counts.columns = ['Nationality', 'Count']
    top_nationality_pie_chart = px.pie(top_nationality_counts.head(15), names='Nationality', values='Count',
                                       title="Top Nationalities", hole=0.4)

    # Pie chart for top paid countries distribution
    top_paid_countries = df[df['Status'] == 'Paid']['Nationality'].value_counts().reset_index()
    top_paid_countries.columns = ['Country', 'Count']
    top_paid_countries_pie_chart = px.pie(top_paid_countries.head(15), names='Country', values='Count',
                                          title="Top Paid Countries", hole=0.4)

    # Bar chart for top 6 regions distribution
    region_counts = df['Region'].value_counts().reset_index()
    region_counts.columns = ['Region', 'Count']
    top_regions_bar_chart = px.bar(region_counts.head(6), x='Region', y='Count',
                                   title="Top 6 Regions by Student Count", text='Count')

    # Pie chart for top 10 programs applied
    applied_programs = df[df['Status'].str.strip().str.lower() == 'applied']['Program'].value_counts().reset_index()
    applied_programs.columns = ['Program', 'Count']
    if not applied_programs.empty:
        top_programs_applied_pie_chart = px.pie(applied_programs.head(10), names='Program', values='Count',
                                                title="Top 10 Programs Applied", hole=0.4)
    else:
        top_programs_applied_pie_chart = {}  # Return an empty chart if no data for applied programs

    # Stacked bar chart for top 7 paid programs
    paid_programs = df[df['Status'] == 'Paid'].groupby('Program').size().reset_index(name='Count')
    top_7_paid_programs = paid_programs.nlargest(7, 'Count')
    top_programs_paid_bar_chart = px.bar(
        top_7_paid_programs,
        x='Program',
        y='Count',
        title="Total Paid for Top 7 Programs",
        color='Program',
        text='Count'
    )
    top_programs_paid_bar_chart.update_traces(
        textposition='inside',
        marker=dict(line=dict(width=1, color='black')),
        width=0.6
    )
    top_programs_paid_bar_chart.update_xaxes(tickangle=-75)
    top_programs_paid_bar_chart.update_layout(
        height=600,
        width=620,
        margin=dict(l=40, r=40, t=70, b=150),
        showlegend=False
    )

    # Logic for the top 10 agents based on application numbers
    if 'Created By' in df.columns:
        agent_counts = df['Created By'].value_counts().reset_index()
        agent_counts.columns = ['Agent', 'Application Count']
        top_10_agents = agent_counts.head(10)  # Get top 10 agents by application count

        # Create a bar chart for the top 10 agents
        top_agents_bar_chart = px.bar(
            top_10_agents,
            x='Agent',
            y='Application Count',
            title='Top 10 Agents by Application Number',
            text='Application Count'
        )
        top_agents_bar_chart.update_traces(textposition='inside')
        top_agents_bar_chart.update_layout(
            xaxis_title="Agent",
            yaxis_title="Application Count",
            height=400,
            width=600
        )
    else:
        top_agents_bar_chart = {}  # Empty chart if no agent data

    # Pie chart for top paid regions distribution
    top_paid_regions = df[df['Status'] == 'Paid']['Region'].value_counts().reset_index()
    top_paid_regions.columns = ['Region', 'Count']
    top_paid_regions_pie_chart = px.pie(top_paid_regions.head(10), names='Region', values='Count',
                                        title="Top Paid Regions", hole=0.4)

    # Return all necessary outputs
    return (total_students, total_paid, top_nationality, top_program, top_status, performance_metric, pie_chart,
            top_nationality_pie_chart, top_paid_countries_pie_chart, top_regions_bar_chart,
            top_agents_bar_chart, top_paid_regions_pie_chart, top_programs_applied_pie_chart, top_programs_paid_bar_chart)


# Run the app
if __name__ == '__main__':
    app.run_server(debug=True)
