import dash
from dash import dcc, html, Input, Output, dash_table
import pandas as pd
import plotly.express as px
import math
# โหลดไฟล์ Excel
df = pd.read_excel("AJ_Thana\myTCAS\mytcas-dashboard\coe_and_aie_with_major.xlsx")

# เพิ่มคอลัมน์ชื่อสาขา
df["ชื่อสาขา"] = df["ชื่อหลักสูตร"].apply(
    lambda x: "วิศวกรรมปัญญาประดิษฐ์" if "ปัญญาประดิษฐ์" in str(x) else "วิศวกรรมคอมพิวเตอร์"
)

# แปลงค่าใช้จ่ายเป็นตัวเลข
df['ค่าใช้จ่าย'] = pd.to_numeric(df['ค่าใช้จ่าย'], errors='coerce')

min_fee = int(df['ค่าใช้จ่าย'].min())
true_max_fee = int(df['ค่าใช้จ่าย'].max())

step = 1000
max_fee = int(math.ceil(true_max_fee / step) * step)

marks = {
    min_fee: {
        'label': f"{min_fee:,}",
        'style': {'color': 'white'}
    },
    true_max_fee: {
        'label': f"{true_max_fee:,}",
        'style': {'color': 'white'}
    }
}


if true_max_fee not in marks:
    marks[true_max_fee] = {'label': f"{true_max_fee:,}", 'style': {'color': 'white'}}

app = dash.Dash(__name__)
app.title = "MyTCAS Dashboard"

# ใส่ Google Font
app.css.append_css({
    'external_url': 'https://fonts.googleapis.com/css2?family=Poppins:wght@400;600&display=swap'
})

# Style constants
BASE_FONT = "'Poppins', sans-serif"
PRIMARY_COLOR = "#FFFFFF"
ACCENT_COLOR = "#002443"
BG_GRADIENT = "linear-gradient(135deg, #002443 100%, #002443 100%)"

LABEL_STYLE = {
    'fontWeight': '450',
    'marginBottom': '8px',
    'fontSize': '16px',
    'color': PRIMARY_COLOR,
    'fontFamily': BASE_FONT,
    # 'marginTop': '10px',
    # 'marginBottom': '10px',
}

DROPDOWN_STYLE = {
    'borderRadius': '8px',
    'boxShadow': '0 2px 5px rgba(0,0,0,0.15)',
    'transition': 'box-shadow 0.3s ease',
    'fontFamily': BASE_FONT,
    'marginTop': '6px',
    'marginBottom': '20px',
}

H2_STYLE = {
    'textAlign': 'center',
    'color': PRIMARY_COLOR,
    'fontWeight': '600',
    'fontFamily': BASE_FONT,
    'fontSize': '2.5rem',
    'userSelect': 'none',
    'marginTop': '40px',
    'textShadow': '1px 1px 2px rgba(0,0,0,0.1)'
}

# Layout
app.layout = html.Div([
    html.Div([
        html.H1("MyTCAS Dashboard", style={
            'fontFamily': "'Poppins', 'Segoe UI', sans-serif",
            'fontWeight': '600',
            'fontSize': '36px',
            'color': "#002443",
            'letterSpacing': '1px',
            'textShadow': '1px 1px 2px rgba(0,0,0,0.1)',
            'textAlign': 'center',
            'margin': '0',
            "backgroundColor": "#FFFFFF",
        })
    ], style={
        'background': BG_GRADIENT,
        'padding': '15px 0',  
        'boxShadow': '0 2px 6px rgba(0,0,0,0.1)',
        'marginBottom': '20px', 
    }),



    html.Div([
        html.Div([
            html.Label("มหาวิทยาลัย:", style=LABEL_STYLE),
            dcc.Dropdown(
                id='university-dropdown',
                options=[{'label': u, 'value': u} for u in sorted(df['มหาวิทยาลัย'].unique())],
                placeholder="เลือกมหาวิทยาลัย",
                clearable=True,
                style=DROPDOWN_STYLE,
            ),
            html.Label("ภูมิภาค:", style={**LABEL_STYLE, 'marginTop': '24px'}),
            dcc.Dropdown(
                id='region-dropdown',
                placeholder="เลือกภูมิภาค",
                multi=True,
                clearable=True,
                style=DROPDOWN_STYLE,
            ),
        ], style={'width': '48%', 'display': 'inline-block', 'verticalAlign': 'top'}),

        html.Div([
            html.Label("หลักสูตร (ภาษาไทย/นานาชาติ):", style=LABEL_STYLE),
            dcc.Dropdown(
                id='course-dropdown',
                placeholder="เลือกหลักสูตร",
                multi=True,
                clearable=True,
                style=DROPDOWN_STYLE,
            ),
            html.Label("หลักสูตร (CoE/AIE):", style={**LABEL_STYLE, 'marginTop': '24px'}),
            dcc.Dropdown(
                id='major-dropdown',
                options=[{'label': m, 'value': m} for m in sorted(df['ชื่อสาขา'].unique())],
                placeholder="เลือกหลักสูตร",
                multi=True,
                clearable=True,
                style=DROPDOWN_STYLE,
            ),
        ], style={'width': '48%', 'display': 'inline-block', 'verticalAlign': 'top'}),
    ], style={'width': '90%', 'margin': '0 auto', 'display': 'flex', 'justifyContent': 'space-between'}),

    html.Div([
        html.Label("ค่าใช้จ่าย (บาท):", style=LABEL_STYLE),
            dcc.RangeSlider(
                id='fee-slider',
                min=min_fee,
                max=max_fee,
                step=step,
                value=[min_fee, max_fee],
                marks=marks,  # ✅ ใช้เฉพาะ min/max
                tooltip={"placement": "bottom"},
                updatemode='mouseup',
            ),


    ], style={'width': '90%', 'margin': '10px auto 40px auto'}),

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
                'color': "#416181"
            },
            style_header={
                'backgroundColor': PRIMARY_COLOR,
                'color': "#000000",
                'fontWeight': '450',
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

    # กรองมหาวิทยาลัย ถ้ามีการเลือก
    if selected_uni:
        filtered = filtered[filtered['มหาวิทยาลัย'] == selected_uni]

    # กรองภูมิภาค ถ้ามีการเลือก และแปลงให้เป็น list เสมอ
    if selected_regions:
        if isinstance(selected_regions, str):
            selected_regions = [selected_regions]
        filtered = filtered[filtered['ภาค'].isin(selected_regions)]

    # กรองหลักสูตร ถ้ามีการเลือก และแปลงเป็น list เสมอ
    if selected_courses:
        if isinstance(selected_courses, str):
            selected_courses = [selected_courses]
        filtered = filtered[filtered['หลักสูตร'].isin(selected_courses)]

    # กรองชื่อสาขา ถ้ามีการเลือก และแปลงเป็น list เสมอ
    if selected_majors:
        if isinstance(selected_majors, str):
            selected_majors = [selected_majors]
        filtered = filtered[filtered['ชื่อสาขา'].isin(selected_majors)]

    # กรองช่วงค่าใช้จ่าย
    # เช็คกรณีค่าใช้จ่ายเป็นตัวเลขจริง (int/float)
    filtered = filtered[
        (filtered['ค่าใช้จ่าย'] >= fee_range[0]) & (filtered['ค่าใช้จ่าย'] <= fee_range[1])
    ]

    # สร้างกราฟวงกลมของภูมิภาค
    if not filtered.empty:
        region_pie = px.pie(
            filtered,
            names='ภาค',
            title="สัดส่วนหลักสูตรตามภาคของประเทศ",
            hole=0.4,
            color_discrete_sequence=px.colors.sequential.Plasma_r
        )
    else:
        region_pie = px.pie(title="ไม่มีข้อมูล")

    # สร้างกราฟแท่งนับจำนวนหลักสูตร
    type_counts = filtered['หลักสูตร'].value_counts().reset_index()
    type_counts.columns = ['หลักสูตร', 'จำนวน']

    if not type_counts.empty:
        type_bar = px.bar(
            type_counts,
            x='หลักสูตร', y='จำนวน',
            labels={'หลักสูตร': 'ประเภทหลักสูตร', 'จำนวน': 'จำนวน'},
            title="จำนวนหลักสูตรทั่วไปกับนานาชาติ",
            color='หลักสูตร',
            color_discrete_sequence=px.colors.qualitative.Pastel
        )
    else:
        type_bar = px.bar(title="ไม่มีข้อมูล")

    # ส่งข้อมูลตาราง
    table_data = filtered.to_dict('records')

    return region_pie, type_bar, table_data



if __name__ == '__main__':
    app.run(debug=True)
