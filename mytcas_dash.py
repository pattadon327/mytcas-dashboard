import dash
from dash import dcc, html, Input, Output, dash_table
import pandas as pd
import plotly.express as px

# โหลดไฟล์ Excel
df = pd.read_excel("AJ_Thana\coe_and_aie_with_major.xlsx")

# เพิ่มคอลัมน์ชื่อสาขา
df["ชื่อสาขา"] = df["ชื่อหลักสูตร"].apply(
    lambda x: "วิศวกรรมปัญญาประดิษฐ์" if "ปัญญาประดิษฐ์" in str(x) else "วิศวกรรมคอมพิวเตอร์"
)

# หา min/max ค่าใช้จ่าย
min_fee, max_fee = df['ค่าใช้จ่าย'].min(), df['ค่าใช้จ่าย'].max()

app = dash.Dash(__name__)
app.title = "Dashboard หลักสูตรมหาวิทยาลัย"

# ใส่ Google Font
app.css.append_css({
    'external_url': 'https://fonts.googleapis.com/css2?family=Poppins:wght@400;600&display=swap'
})

# Style constants
BASE_FONT = "'Poppins', sans-serif"
PRIMARY_COLOR = "#083A17"
ACCENT_COLOR = "#ff7f0e"
BG_GRADIENT = "linear-gradient(135deg, #e0f7fa 0%, #80deea 100%)"

LABEL_STYLE = {
    'fontWeight': '600',
    'marginBottom': '8px',
    'fontSize': '16px',
    'color': PRIMARY_COLOR,
    'fontFamily': BASE_FONT,
}

DROPDOWN_STYLE = {
    'borderRadius': '8px',
    'boxShadow': '0 2px 5px rgba(0,0,0,0.15)',
    'transition': 'box-shadow 0.3s ease',
    'fontFamily': BASE_FONT,
}

H2_STYLE = {
    'textAlign': 'center',
    'color': PRIMARY_COLOR,
    'fontWeight': '700',
    'fontFamily': BASE_FONT,
    'fontSize': '2.5rem',
    'userSelect': 'none',
    'marginTop': '40px',
    'textShadow': '1px 1px 2px rgba(0,0,0,0.1)'
}

# Layout
app.layout = html.Div([
    html.Div([
        html.H1("🎓 Dashboard หลักสูตรมหาวิทยาลัย", style=H2_STYLE)
    ], style={
        'background': BG_GRADIENT,
        'padding': '30px 0',
        'boxShadow': '0 4px 10px rgba(0,0,0,0.1)',
        'marginBottom': '40px',
    }),

    html.Div([
        html.Div([
            html.Label("🏫 เลือกมหาวิทยาลัย:", style=LABEL_STYLE),
            dcc.Dropdown(
                id='university-dropdown',
                options=[{'label': u, 'value': u} for u in sorted(df['มหาวิทยาลัย'].unique())],
                placeholder="เลือกมหาวิทยาลัย (หรือเว้นไว้)...",
                clearable=True,
                style=DROPDOWN_STYLE,
            ),
            html.Label("📍 เลือกภาค:", style={**LABEL_STYLE, 'marginTop': '24px'}),
            dcc.Dropdown(
                id='region-dropdown',
                placeholder="เลือกภาค (หรือเว้นไว้)...",
                multi=True,
                clearable=True,
                style=DROPDOWN_STYLE,
            ),
        ], style={'width': '48%', 'display': 'inline-block', 'verticalAlign': 'top'}),

        html.Div([
            html.Label("🎓 เลือกหลักสูตร:", style=LABEL_STYLE),
            dcc.Dropdown(
                id='course-dropdown',
                placeholder="เลือกหลักสูตร (หรือเว้นไว้)...",
                multi=True,
                clearable=True,
                style=DROPDOWN_STYLE,
            ),
            html.Label("👨‍💻 เลือกชื่อสาขา:", style={**LABEL_STYLE, 'marginTop': '24px'}),
            dcc.Dropdown(
                id='major-dropdown',
                options=[{'label': m, 'value': m} for m in sorted(df['ชื่อสาขา'].unique())],
                placeholder="เลือกชื่อสาขา (หรือเว้นไว้)...",
                multi=True,
                clearable=True,
                style=DROPDOWN_STYLE,
            ),
        ], style={'width': '48%', 'display': 'inline-block', 'verticalAlign': 'top'}),
    ], style={'width': '90%', 'margin': '0 auto', 'display': 'flex', 'justifyContent': 'space-between'}),

    html.Div([
        html.Label("💰 เลือกช่วงค่าใช้จ่าย (บาท):", style=LABEL_STYLE),
        dcc.RangeSlider(
            id='fee-slider',
            min=min_fee, max=max_fee,
            step=1000,
            value=[min_fee, max_fee],
            marks={int(f): str(int(f)) for f in range(int(min_fee), int(max_fee)+1, 10000)},
            tooltip={"placement": "bottom"},
            updatemode='mouseup',
        ),
    ], style={'width': '90%', 'margin': '40px auto 60px auto'}),

    html.Div([
        dcc.Graph(id='region-pie', style={
            'width': '48%',
            'display': 'inline-block',
            'boxShadow': '0 4px 12px rgba(0,0,0,0.1)',
            'borderRadius': '12px',
            'padding': '10px',
            'backgroundColor': 'white',
        }),
        dcc.Graph(id='type-bar', style={
            'width': '48%',
            'display': 'inline-block',
            'boxShadow': '0 4px 12px rgba(0,0,0,0.1)',
            'borderRadius': '12px',
            'padding': '10px',
            'backgroundColor': 'white',
        }),
    ], style={'width': '90%', 'margin': 'auto'}),

    html.Div([
        dash_table.DataTable(
            id='table',
            columns=[{"name": i, "id": i} for i in df.columns],
            data=df.to_dict('records'),
            page_size=10,
            style_table={'overflowX': 'auto'},
            style_cell={
                'textAlign': 'left',
                'padding': '8px',
                'fontFamily': BASE_FONT,
                'fontSize': '14px',
                'color': '#34495e'
            },
            style_header={
                'backgroundColor': PRIMARY_COLOR,
                'color': 'white',
                'fontWeight': '600',
            },
            style_data_conditional=[
                {'if': {'row_index': 'odd'}, 'backgroundColor': '#ecf0f1'}
            ],
            fixed_rows={'headers': True},
            filter_action='native',
            sort_action='native',
            page_action='native',
        )
    ], style={
        'width': '90%',
        'margin': '60px auto',
        'backgroundColor': 'white',
        'padding': '20px',
        'borderRadius': '12px',
        'boxShadow': '0 8px 20px rgba(0,0,0,0.12)',
    }),

], style={
    'background': BG_GRADIENT,
    'minHeight': '100vh',
    'paddingBottom': '60px',
})


@app.callback(
    [Output('region-dropdown', 'options'),
     Output('course-dropdown', 'options')],
    [Input('university-dropdown', 'value'),
     Input('fee-slider', 'value')]
)
def update_region_course_options(selected_uni, fee_range):
    filtered = df.copy()
    if selected_uni:
        filtered = filtered[filtered['มหาวิทยาลัย'] == selected_uni]

    filtered = filtered[
        (filtered['ค่าใช้จ่าย'] >= fee_range[0]) & (filtered['ค่าใช้จ่าย'] <= fee_range[1])
    ]

    region_options = [{'label': r, 'value': r} for r in sorted(filtered['ภาค'].dropna().unique())]
    course_options = [{'label': c, 'value': c} for c in sorted(filtered['หลักสูตร'].dropna().unique())]

    return region_options, course_options


@app.callback(
    [Output('region-pie', 'figure'),
     Output('type-bar', 'figure'),
     Output('table', 'data')],
    [Input('university-dropdown', 'value'),
     Input('region-dropdown', 'value'),
     Input('course-dropdown', 'value'),
     Input('major-dropdown', 'value'),
     Input('fee-slider', 'value')]
)
def update_graphs_and_table(selected_uni, selected_regions, selected_courses, selected_majors, fee_range):
    filtered = df.copy()

    if selected_uni:
        filtered = filtered[filtered['มหาวิทยาลัย'] == selected_uni]

    if selected_regions:
        if isinstance(selected_regions, str):
            selected_regions = [selected_regions]
        filtered = filtered[filtered['ภาค'].isin(selected_regions)]

    if selected_courses:
        if isinstance(selected_courses, str):
            selected_courses = [selected_courses]
        filtered = filtered[filtered['หลักสูตร'].isin(selected_courses)]

    if selected_majors:
        if isinstance(selected_majors, str):
            selected_majors = [selected_majors]
        filtered = filtered[filtered['ชื่อสาขา'].isin(selected_majors)]

    filtered = filtered[
        (filtered['ค่าใช้จ่าย'] >= fee_range[0]) & (filtered['ค่าใช้จ่าย'] <= fee_range[1])
    ]

    if not filtered.empty:
        region_pie = px.pie(
            filtered,
            names='ภาค',
            title="📍 สัดส่วนหลักสูตรตามภาคของประเทศ",
            hole=0.4,
            color_discrete_sequence=px.colors.sequential.Plasma_r
        )
    else:
        region_pie = px.pie(title="ไม่มีข้อมูล")

    type_counts = filtered['หลักสูตร'].value_counts().reset_index()
    type_counts.columns = ['หลักสูตร', 'จำนวน']

    if not type_counts.empty:
        type_bar = px.bar(
            type_counts,
            x='หลักสูตร', y='จำนวน',
            labels={'หลักสูตร': 'ประเภทหลักสูตร', 'จำนวน': 'จำนวน'},
            title="🏷️ จำนวนหลักสูตรทั่วไป vs. นานาชาติ",
            color='หลักสูตร',
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
    else:
        type_bar = px.bar(title="ไม่มีข้อมูล")

    table_data = filtered.to_dict('records')

    return region_pie, type_bar, table_data


if __name__ == '__main__':
    app.run(debug=True)
