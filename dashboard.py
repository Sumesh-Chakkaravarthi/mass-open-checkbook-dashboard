#!/usr/bin/env python3
"""
ALY 6980 Capstone — Interactive Dashboard
Massachusetts Open Checkbook: Vendor Contract & SDO Commitment Analysis

A Plotly Dash interactive dashboard with 4 tabs, interactive filters,
KPI cards, and additional dashboard-exclusive charts.

Usage:
    python dashboard.py
    Open http://127.0.0.1:8050 in your browser.

Author: Sumesh Chakkaravarthi
Date: February 2026
"""

import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from dash import Dash, html, dcc, Input, Output, callback
import textwrap

# ──────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────
VENDOR_FILE = "Copy of Vendor Contact Details (1).xlsx"
CATEGORIZED_FILE = "List of Categorized_Companies (1).xlsx"

CATEGORY_NAMES = {
    'ITE': 'IT Equipment & Services',
    'ITS': 'IT Software & Services',
    'ITT': 'Telecom & Networking',
    'FAC': 'Facilities General',
    'VEH': 'Vehicle Acquisition & Maint.',
    'GRO': 'Food & Food Service',
    'LND': 'Facility Landscaping',
    'MED': 'Health & Medical',
    'MRO': 'Maintenance, Repair & Ops',
    'OFF': 'Office Supplies',
    'PRF': 'Professional Services',
    'PSE': 'Public Safety & Security',
    'SFC': 'Sustainable Facilities',
    'TRD': 'Tradespersons',
    'WMR': 'Waste Mgmt & Recycling',
}

SKIP_SHEETS = ['Abbreviations ']

# Color scheme
COLORS = {
    'bg': '#0F172A',
    'card': '#1E293B',
    'card_border': '#334155',
    'accent': '#06B6D4',
    'accent2': '#8B5CF6',
    'accent3': '#F43F5E',
    'accent4': '#10B981',
    'text': '#F1F5F9',
    'text_muted': '#94A3B8',
    'grid': '#334155',
}

CHART_TEMPLATE = 'plotly_dark'

# ──────────────────────────────────────────────────────────────────
# Data Loading (same as EDA, but streamlined)
# ──────────────────────────────────────────────────────────────────
def load_vendor_data(filepath):
    xl = pd.ExcelFile(filepath)
    frames = []
    for sheet in xl.sheet_names:
        if sheet.strip() in [s.strip() for s in SKIP_SHEETS]:
            continue
        df = pd.read_excel(filepath, sheet_name=sheet, header=0)
        df = df.iloc[:, :7]
        df.columns = ['Contract_Code', 'Name', 'Company', 'Role', 'Email', 'Phone', 'SDO_Pct']
        df['Category'] = sheet.strip()

        metadata_kw = ['Master Contract', 'Solicitation Enabled', 'Master MBPO',
                        'Bid and Contract', 'Category and Vendor', 'Category Development',
                        'Mass Gov', 'OSD Help Desk', 'N/A']
        mask = df['Company'].apply(
            lambda x: not any(kw.lower() in str(x).lower() for kw in metadata_kw)
        )
        df = df[mask].copy()
        df['SDO_Pct'] = pd.to_numeric(df['SDO_Pct'], errors='coerce')
        df['Company'] = df['Company'].astype(str).str.replace('\n', ' ', regex=False).str.strip()
        df['Name'] = df['Name'].astype(str).str.replace('\n', ' ', regex=False).str.strip()
        df['Contract_Code'] = df['Contract_Code'].astype(str).str.replace('\n', ' ', regex=False).str.strip()
        df['Category_Label'] = df['Category'].map(CATEGORY_NAMES)
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def load_categorized_companies(filepath):
    xl = pd.ExcelFile(filepath)
    frames = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(filepath, sheet_name=sheet, header=None)
        industry = sheet.strip()
        for col_idx, typ in [(1, 'National & Local'), (2, 'Local'), (3, 'SGC Target')]:
            companies = df.iloc[2:, col_idx].dropna().astype(str).str.strip().tolist()
            for c in companies:
                if c and c != 'nan':
                    frames.append({'Industry': industry, 'Company': c, 'Type': typ})
    return pd.DataFrame(frames)


# ──────────────────────────────────────────────────────────────────
# Load data at module level
# ──────────────────────────────────────────────────────────────────
print("Loading data...")
vendor_df = load_vendor_data(VENDOR_FILE)
cat_df = load_categorized_companies(CATEGORIZED_FILE)
print(f"✅ {len(vendor_df):,} vendor records, {len(cat_df):,} categorized companies")

# Pre-compute common derived datasets
vendor_with_sdo = vendor_df.dropna(subset=['SDO_Pct']).copy()
vendor_with_sdo = vendor_with_sdo[vendor_with_sdo['SDO_Pct'] > 0]
vendor_with_sdo['SDO_Capped'] = vendor_with_sdo['SDO_Pct'].clip(upper=1.0)


# ──────────────────────────────────────────────────────────────────
# Helper: KPI card HTML
# ──────────────────────────────────────────────────────────────────
def kpi_card(title, value, subtitle="", color=COLORS['accent']):
    return html.Div(
        children=[
            html.H4(title, style={'color': COLORS['text_muted'], 'margin': '0 0 8px 0',
                                   'fontSize': '13px', 'textTransform': 'uppercase',
                                   'letterSpacing': '1px', 'fontWeight': '500'}),
            html.H2(value, style={'color': color, 'margin': '0 0 4px 0',
                                   'fontSize': '32px', 'fontWeight': '700'}),
            html.P(subtitle, style={'color': COLORS['text_muted'], 'margin': '0',
                                     'fontSize': '12px'}),
        ],
        style={
            'background': COLORS['card'],
            'borderRadius': '12px',
            'padding': '24px',
            'border': f"1px solid {COLORS['card_border']}",
            'flex': '1',
            'minWidth': '200px',
            'textAlign': 'center',
        }
    )


# ──────────────────────────────────────────────────────────────────
# Chart layout helper
# ──────────────────────────────────────────────────────────────────
def chart_layout(fig, title="", height=500):
    fig.update_layout(
        template=CHART_TEMPLATE,
        title=dict(text=title, font=dict(size=16, color=COLORS['text'])),
        paper_bgcolor=COLORS['card'],
        plot_bgcolor=COLORS['card'],
        font=dict(color=COLORS['text'], family='Inter, system-ui, sans-serif'),
        margin=dict(l=60, r=30, t=60, b=60),
        height=height,
        legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(size=11)),
    )
    fig.update_xaxes(gridcolor=COLORS['grid'], zeroline=False)
    fig.update_yaxes(gridcolor=COLORS['grid'], zeroline=False)
    return fig


# ══════════════════════════════════════════════════════════════════
# CHART GENERATORS
# ══════════════════════════════════════════════════════════════════

# ── BQ1: Top IT Companies by SDO ──
def make_bq1_chart(it_filter='All'):
    it_cats = ['ITE', 'ITS', 'ITT'] if it_filter == 'All' else [it_filter]
    df = vendor_with_sdo[vendor_with_sdo['Category'].isin(it_cats)].copy()
    top = df.groupby(['Company', 'Category_Label']).agg(
        SDO_Max=('SDO_Capped', 'max')
    ).reset_index().sort_values('SDO_Max', ascending=True).tail(15)

    fig = px.bar(top, y='Company', x='SDO_Max', color='Category_Label',
                 orientation='h', color_discrete_sequence=[COLORS['accent'], COLORS['accent2'], COLORS['accent3']])
    fig.update_traces(texttemplate='%{x:.0%}', textposition='outside')
    fig.update_layout(xaxis_tickformat='.0%', xaxis_title='SDO Commitment %',
                      yaxis_title='', legend_title='')
    return chart_layout(fig, f'BQ1: Top 15 IT Companies by SDO Commitment ({it_filter})', 550)


# ── BQ2: Average SDO by Category ──
def make_bq2_chart():
    df = vendor_with_sdo.copy()
    stats = df.groupby('Category_Label').agg(
        Avg=('SDO_Capped', 'mean'), Count=('SDO_Capped', 'count')
    ).reset_index().sort_values('Avg', ascending=True)

    fig = px.bar(stats, y='Category_Label', x='Avg', orientation='h',
                 color='Avg', color_continuous_scale='Viridis',
                 hover_data={'Count': True})
    fig.update_traces(texttemplate='%{x:.1%}', textposition='outside')
    fig.update_layout(xaxis_tickformat='.0%', xaxis_title='Average SDO %',
                      yaxis_title='', coloraxis_colorbar_title='Avg SDO')
    return chart_layout(fig, 'BQ2: Average SDO Commitment by Procurement Category', 550)


# ── BQ3: Vendor Distribution Treemap ──
def make_bq3_chart():
    counts = vendor_df.groupby('Category_Label')['Company'].nunique().reset_index()
    counts.columns = ['Category', 'Vendors']
    fig = px.treemap(counts, path=['Category'], values='Vendors',
                     color='Vendors', color_continuous_scale='Teal')
    fig.update_traces(textinfo='label+value+percent root',
                      textfont=dict(size=13))
    return chart_layout(fig, 'BQ3: Vendor Distribution Across Procurement Categories', 500)


# ── BQ4: Contract Sub-Categories ──
def make_bq4_chart():
    codes = vendor_df.groupby('Contract_Code')['Company'].nunique()
    codes = codes[codes.index.str.len() <= 15].sort_values(ascending=False).head(20)
    df = codes.reset_index()
    df.columns = ['Code', 'Vendors']

    fig = px.bar(df, x='Code', y='Vendors', color='Vendors',
                 color_continuous_scale='Sunset')
    fig.update_traces(texttemplate='%{y}', textposition='outside')
    fig.update_layout(xaxis_title='Contract Code', yaxis_title='Unique Vendors')
    return chart_layout(fig, 'BQ4: Top 20 Contract Sub-Categories by Vendor Count', 450)


# ── BQ5: National vs Local ──
def make_bq5_chart():
    pivot = cat_df.groupby(['Industry', 'Type']).size().reset_index(name='Count')
    fig = px.bar(pivot, y='Industry', x='Count', color='Type', orientation='h',
                 barmode='group',
                 color_discrete_map={
                     'National & Local': COLORS['accent'],
                     'Local': COLORS['accent2'],
                     'SGC Target': COLORS['accent3']
                 })
    fig.update_layout(yaxis_title='', xaxis_title='Number of Companies', legend_title='')
    return chart_layout(fig, 'BQ5: National vs Local Company Presence by Industry', 500)


# ── BQ6: SDO Distribution Box Plot ──
def make_bq6_chart():
    cat_counts = vendor_with_sdo.groupby('Category').size()
    valid_cats = cat_counts[cat_counts >= 10].index.tolist()
    df = vendor_with_sdo[vendor_with_sdo['Category'].isin(valid_cats)].copy()

    fig = px.box(df, x='Category_Label', y='SDO_Capped', color='Category_Label',
                 color_discrete_sequence=px.colors.qualitative.Set2)
    fig.update_layout(xaxis_title='', yaxis_title='SDO Commitment %',
                      yaxis_tickformat='.0%', showlegend=False)
    fig.update_xaxes(tickangle=-45)
    return chart_layout(fig, 'BQ6: SDO Commitment Distribution & Outliers by Category', 500)


# ── BQ7: SDO Coverage Rate ──
def make_bq7_chart():
    coverage = vendor_df.groupby('Category_Label').apply(
        lambda g: pd.Series({
            'Has SDO': g['SDO_Pct'].notna().sum(),
            'No SDO': g['SDO_Pct'].isna().sum(),
        })
    ).reset_index()
    melted = coverage.melt(id_vars='Category_Label', var_name='Status', value_name='Count')

    fig = px.bar(melted, y='Category_Label', x='Count', color='Status', orientation='h',
                 barmode='group',
                 color_discrete_map={'Has SDO': COLORS['accent4'], 'No SDO': COLORS['accent3']})
    fig.update_layout(yaxis_title='', xaxis_title='Number of Vendors', legend_title='')
    return chart_layout(fig, 'BQ7: SDO Coverage – Vendors with Valid SDO vs Missing', 550)


# ── BQ8: Scatter — Vendor Count vs Avg SDO ──
def make_bq8_chart():
    df = vendor_with_sdo[vendor_with_sdo['SDO_Capped'] <= 1.0].copy()
    stats = df.groupby('Category').agg(
        Vendors=('Company', 'nunique'),
        Avg_SDO=('SDO_Capped', 'mean')
    ).reset_index()
    stats['Label'] = stats['Category'].map(CATEGORY_NAMES)

    fig = px.scatter(stats, x='Vendors', y='Avg_SDO', text='Category',
                     size='Vendors', color='Avg_SDO',
                     color_continuous_scale='Viridis', size_max=40)
    fig.update_traces(textposition='top center', textfont=dict(size=11, color=COLORS['text']))
    fig.update_layout(xaxis_title='Number of Unique Vendors',
                      yaxis_title='Average SDO %', yaxis_tickformat='.0%')

    # Add trendline
    from scipy import stats as sp_stats
    if len(stats) > 2:
        slope, intercept, r, p, se = sp_stats.linregress(stats['Vendors'], stats['Avg_SDO'])
        x_range = np.linspace(stats['Vendors'].min(), stats['Vendors'].max(), 50)
        fig.add_trace(go.Scatter(x=x_range, y=intercept + slope * x_range,
                                  mode='lines', name=f'Trend (r={r:.2f})',
                                  line=dict(color=COLORS['accent3'], dash='dash', width=2)))
    return chart_layout(fig, 'BQ8: Vendor Count vs Average SDO Commitment', 500)


# ── BQ9: Industry Diversity Heatmap ──
def make_bq9_chart():
    pivot = cat_df.groupby(['Industry', 'Type']).size().unstack(fill_value=0)
    fig = px.imshow(pivot, text_auto=True, color_continuous_scale='YlOrRd',
                    labels=dict(x='Company Type', y='Industry', color='Count'))
    fig.update_layout(xaxis_title='', yaxis_title='')
    return chart_layout(fig, 'BQ9: Industry Diversity – National vs Local vs SGC Target', 500)


# ── BQ10: IT Vendor Concentration Donut ──
def make_bq10_chart():
    it_df = vendor_df[vendor_df['Category'].isin(['ITE', 'ITS', 'ITT'])]
    company_counts = it_df.groupby('Company')['Contract_Code'].nunique().sort_values(ascending=False)
    top10 = company_counts.head(10)
    rest = company_counts.iloc[10:].sum()

    labels = list(top10.index) + [f'Other ({len(company_counts)-10} companies)']
    values = list(top10.values) + [rest]

    fig = go.Figure(go.Pie(labels=labels, values=values, hole=0.5,
                           textinfo='label+percent', textposition='outside',
                           marker=dict(colors=px.colors.qualitative.Set3)))
    fig.update_layout(showlegend=False)
    return chart_layout(fig, 'BQ10: IT Sector Vendor Concentration', 500)


# ── Dashboard-Only: SDO Histogram ──
def make_sdo_histogram():
    fig = px.histogram(vendor_with_sdo, x='SDO_Capped', nbins=50,
                       color_discrete_sequence=[COLORS['accent']],
                       labels={'SDO_Capped': 'SDO Commitment %'})
    fig.update_layout(xaxis_tickformat='.0%', xaxis_title='SDO Commitment %',
                      yaxis_title='Number of Vendors', bargap=0.05)
    return chart_layout(fig, 'Overall SDO Commitment Distribution (All Categories)', 400)


# ── Dashboard-Only: Radar Chart — IT Sub-Categories ──
def make_it_radar():
    metrics = []
    for cat in ['ITE', 'ITS', 'ITT']:
        df_cat = vendor_with_sdo[vendor_with_sdo['Category'] == cat]
        metrics.append({
            'Category': CATEGORY_NAMES[cat],
            'Avg SDO': df_cat['SDO_Capped'].mean() * 100,
            'Median SDO': df_cat['SDO_Capped'].median() * 100,
            'Max SDO': df_cat['SDO_Capped'].max() * 100,
            'Vendor Count': len(df_cat),
            'Coverage Rate': (vendor_df[vendor_df['Category'] == cat]['SDO_Pct'].notna().mean()) * 100,
        })

    categories_radar = ['Avg SDO', 'Median SDO', 'Max SDO', 'Vendor Count', 'Coverage Rate']

    fig = go.Figure()
    colors_radar = [COLORS['accent'], COLORS['accent2'], COLORS['accent3']]
    for i, m in enumerate(metrics):
        # Normalize to 0-100 scale
        vals = [m[c] for c in categories_radar]
        max_vals = [max(me[c] for me in metrics) for c in categories_radar]
        normalized = [(v / mx * 100 if mx > 0 else 0) for v, mx in zip(vals, max_vals)]
        normalized.append(normalized[0])  # close the polygon

        fig.add_trace(go.Scatterpolar(
            r=normalized,
            theta=categories_radar + [categories_radar[0]],
            fill='toself',
            name=m['Category'],
            line=dict(color=colors_radar[i], width=2),
            fillcolor=colors_radar[i].replace(')', ', 0.15)').replace('rgb', 'rgba') if 'rgb' in colors_radar[i] else None,
            opacity=0.8
        ))

    fig.update_layout(
        polar=dict(
            radialaxis=dict(visible=True, range=[0, 110], gridcolor=COLORS['grid'],
                            tickfont=dict(size=9, color=COLORS['text_muted'])),
            angularaxis=dict(gridcolor=COLORS['grid'],
                             tickfont=dict(size=11, color=COLORS['text'])),
            bgcolor=COLORS['card'],
        ),
        showlegend=True,
    )
    return chart_layout(fig, 'IT Sub-Category Comparison (Normalized Radar)', 450)


# ══════════════════════════════════════════════════════════════════
# DASH APP
# ══════════════════════════════════════════════════════════════════
app = Dash(__name__, title="MA Open Checkbook Dashboard")

# CSS
app.index_string = '''
<!DOCTYPE html>
<html>
<head>
    {%metas%}
    <title>{%title%}</title>
    {%favicon%}
    {%css%}
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            background: ''' + COLORS['bg'] + ''';
            color: ''' + COLORS['text'] + ''';
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
        }
        .dash-tab {
            background: transparent !important;
            color: ''' + COLORS['text_muted'] + ''' !important;
            border: none !important;
            padding: 12px 24px !important;
            font-size: 14px !important;
            font-weight: 500 !important;
            text-transform: uppercase !important;
            letter-spacing: 0.5px !important;
        }
        .dash-tab--selected {
            color: ''' + COLORS['accent'] + ''' !important;
            border-bottom: 3px solid ''' + COLORS['accent'] + ''' !important;
        }
        .tab-content { padding: 24px; }
        .chart-card {
            background: ''' + COLORS['card'] + ''';
            border-radius: 12px;
            padding: 16px;
            margin-bottom: 20px;
            border: 1px solid ''' + COLORS['card_border'] + ''';
        }
    </style>
</head>
<body>
    {%app_entry%}
    <footer>
        {%config%}
        {%scripts%}
        {%renderer%}
    </footer>
</body>
</html>
'''

# ── KPI Calculations ──
total_vendors = vendor_df['Company'].nunique()
avg_sdo = vendor_with_sdo['SDO_Capped'].mean()
categories_count = vendor_df['Category'].nunique()
it_vendors = vendor_df[vendor_df['Category'].isin(['ITE', 'ITS', 'ITT'])]['Company'].nunique()
industries_count = cat_df['Industry'].nunique()

# ── Layout ──
app.layout = html.Div([
    # Header
    html.Div([
        html.Div([
            html.H1("Massachusetts Open Checkbook", style={
                'fontSize': '28px', 'fontWeight': '700', 'margin': '0',
                'background': f'linear-gradient(135deg, {COLORS["accent"]}, {COLORS["accent2"]})',
                '-webkit-background-clip': 'text', '-webkit-text-fill-color': 'transparent',
            }),
            html.P("Vendor Contract & SDO Commitment Analysis Dashboard",
                    style={'color': COLORS['text_muted'], 'fontSize': '14px', 'margin': '4px 0 0 0'}),
        ], style={'flex': '1'}),
        html.Div([
            html.Span("ALY 6980 Capstone", style={
                'background': COLORS['card'], 'color': COLORS['accent'],
                'padding': '8px 16px', 'borderRadius': '20px', 'fontSize': '12px',
                'border': f'1px solid {COLORS["card_border"]}', 'fontWeight': '600',
            }),
        ]),
    ], style={
        'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center',
        'padding': '20px 32px', 'borderBottom': f'1px solid {COLORS["card_border"]}',
    }),

    # KPI Row
    html.Div([
        kpi_card("Total Vendors", f"{total_vendors:,}", "Unique companies", COLORS['accent']),
        kpi_card("Average SDO", f"{avg_sdo:.1%}", "Commitment percentage", COLORS['accent2']),
        kpi_card("Categories", str(categories_count), "Procurement categories", COLORS['accent3']),
        kpi_card("IT Vendors", f"{it_vendors:,}", "ITE + ITS + ITT", COLORS['accent4']),
        kpi_card("Industries", str(industries_count), "Company classifications", '#F59E0B'),
    ], style={
        'display': 'flex', 'gap': '16px', 'padding': '24px 32px',
        'overflowX': 'auto',
    }),

    # Tabs
    dcc.Tabs(id='main-tabs', value='tab-1', children=[
        dcc.Tab(label='🖥️ IT Sector SDO', value='tab-1'),
        dcc.Tab(label='📊 Cross-Category', value='tab-2'),
        dcc.Tab(label='✅ Vendor Coverage', value='tab-3'),
        dcc.Tab(label='🏢 Industry Analysis', value='tab-4'),
    ], style={'padding': '0 32px'}),

    # Tab content
    html.Div(id='tab-content', style={'padding': '0 32px 32px 32px'}),

    # Footer
    html.Div([
        html.P("ALY 6980 Capstone Project • Sumesh Chakkaravarthi • February 2026",
               style={'color': COLORS['text_muted'], 'fontSize': '12px', 'textAlign': 'center'}),
    ], style={'padding': '16px', 'borderTop': f'1px solid {COLORS["card_border"]}'}),
], style={'maxWidth': '1400px', 'margin': '0 auto'})


# ── Tab Content Callback ──
@callback(Output('tab-content', 'children'), Input('main-tabs', 'value'))
def render_tab(tab):
    if tab == 'tab-1':
        return html.Div([
            # Filter
            html.Div([
                html.Label("Filter by IT Sub-Category:", style={
                    'color': COLORS['text_muted'], 'fontSize': '13px', 'fontWeight': '500',
                    'marginRight': '12px'
                }),
                dcc.Dropdown(
                    id='it-filter',
                    options=[
                        {'label': 'All IT Categories', 'value': 'All'},
                        {'label': 'ITE — IT Equipment & Services', 'value': 'ITE'},
                        {'label': 'ITS — IT Software & Services', 'value': 'ITS'},
                        {'label': 'ITT — Telecom & Networking', 'value': 'ITT'},
                    ],
                    value='All',
                    clearable=False,
                    style={'width': '300px', 'backgroundColor': COLORS['card'],
                           'color': '#333', 'borderRadius': '8px'},
                ),
            ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '20px', 'marginTop': '20px'}),

            # Charts
            html.Div([
                html.Div([dcc.Graph(id='bq1-chart')], className='chart-card', style={'flex': '1'}),
            ], style={'display': 'flex', 'gap': '20px'}),

            html.Div([
                html.Div([dcc.Graph(figure=make_bq10_chart())], className='chart-card', style={'flex': '1'}),
                html.Div([dcc.Graph(figure=make_it_radar())], className='chart-card', style={'flex': '1'}),
            ], style={'display': 'flex', 'gap': '20px'}),
        ])

    elif tab == 'tab-2':
        return html.Div([
            html.Div([
                html.Div([dcc.Graph(figure=make_bq2_chart())], className='chart-card'),
            ]),
            html.Div([
                html.Div([dcc.Graph(figure=make_bq6_chart())], className='chart-card', style={'flex': '1'}),
                html.Div([dcc.Graph(figure=make_bq8_chart())], className='chart-card', style={'flex': '1'}),
            ], style={'display': 'flex', 'gap': '20px'}),
            html.Div([
                html.Div([dcc.Graph(figure=make_sdo_histogram())], className='chart-card'),
            ]),
        ], style={'paddingTop': '20px'})

    elif tab == 'tab-3':
        return html.Div([
            html.Div([
                html.Div([dcc.Graph(figure=make_bq7_chart())], className='chart-card'),
            ]),
            html.Div([
                html.Div([dcc.Graph(figure=make_bq3_chart())], className='chart-card', style={'flex': '1'}),
                html.Div([dcc.Graph(figure=make_bq4_chart())], className='chart-card', style={'flex': '1'}),
            ], style={'display': 'flex', 'gap': '20px'}),
        ], style={'paddingTop': '20px'})

    elif tab == 'tab-4':
        return html.Div([
            html.Div([
                html.Div([dcc.Graph(figure=make_bq5_chart())], className='chart-card', style={'flex': '1'}),
                html.Div([dcc.Graph(figure=make_bq9_chart())], className='chart-card', style={'flex': '1'}),
            ], style={'display': 'flex', 'gap': '20px'}),
        ], style={'paddingTop': '20px'})


# ── BQ1 dynamic callback ──
@callback(Output('bq1-chart', 'figure'), Input('it-filter', 'value'))
def update_bq1(it_filter):
    return make_bq1_chart(it_filter or 'All')


# ──────────────────────────────────────────────────────────────────
# Run
# ──────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    print("🚀 Starting dashboard at http://127.0.0.1:8050")
    app.run(debug=True, port=8050)
