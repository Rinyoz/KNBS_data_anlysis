import re
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

WORKBOOK_FILE = Path(__file__).with_name('KNBS Data.xlsx')

st.set_page_config(page_title='KNBS KNBS Data Explorer', layout='wide')

st.markdown(
    """
    <style>
    .main { background-color: #f0f2f6; }
    .stMarkdown h1, .stMarkdown h2, .stMarkdown h3 { color: #0f3d69; }
    .metric-container { padding: 12px; border-radius: 12px; background: white; box-shadow: 0 6px 18px rgba(0,0,0,0.08); }
    </style>
    """,
    unsafe_allow_html=True,
)

@st.cache_data
def get_sheet_names():
    if WORKBOOK_FILE.exists():
        return pd.ExcelFile(WORKBOOK_FILE, engine='openpyxl').sheet_names
    return []


def dedupe_columns(columns):
    counts = {}
    result = []
    for col in columns:
        clean_col = col.strip() if isinstance(col, str) else str(col)
        if clean_col in counts:
            counts[clean_col] += 1
            result.append(f"{clean_col}.{counts[clean_col]}")
        else:
            counts[clean_col] = 0
            result.append(clean_col)
    return result


def find_header_row(df):
    for idx, row in df.iterrows():
        non_empty = row.notna().sum()
        if non_empty < 2:
            continue
        string_values = sum(1 for value in row if isinstance(value, str) and value.strip())
        if string_values >= max(1, non_empty // 2):
            return idx
    return 0


def parse_dates(df):
    date_candidates = [
        c for c in df.columns if re.search(r"date|period|month|year|as at|months|arrival", str(c), re.I)
    ]
    if not date_candidates:
        return None

    for candidate in date_candidates:
        try:
            parsed = pd.to_datetime(df[candidate], errors='coerce', dayfirst=True)
            if parsed.notna().sum() > 0:
                df[candidate] = parsed
                return candidate
        except Exception:
            continue
    return None


def normalize_sheet(sheet_name):
    raw = pd.read_excel(WORKBOOK_FILE, sheet_name=sheet_name, header=None, engine='openpyxl')
    raw = raw.loc[:, ~raw.isna().all(axis=0)]
    raw = raw.dropna(how='all').reset_index(drop=True)

    header_idx = find_header_row(raw)
    header = raw.iloc[header_idx].ffill().astype(str).str.strip()
    header = dedupe_columns(header.tolist())

    df = raw.iloc[header_idx + 1 :].copy()
    df.columns = header
    df = df.dropna(how='all').reset_index(drop=True)
    df.columns = [str(col).strip() for col in df.columns]

    for col in df.columns:
        if df[col].dtype == object:
            try:
                df[col] = pd.to_numeric(df[col])
            except (ValueError, TypeError):
                pass

    parse_dates(df)
    return df


def get_data():
    sheet_names = get_sheet_names()
    return {sheet: normalize_sheet(sheet) for sheet in sheet_names}


def find_best_numeric(df, preferred=None):
    if preferred:
        for candidate in preferred:
            if candidate in df.columns and pd.api.types.is_numeric_dtype(df[candidate]):
                return candidate
    numeric = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    return numeric[0] if numeric else None


def show_chart(df, x_col, y_col, title=None, chart_type='line'):
    if chart_type == 'bar':
        fig = px.bar(df, x=x_col, y=y_col, title=title)
    else:
        fig = px.line(df, x=x_col, y=y_col, title=title)
    st.plotly_chart(fig, width='stretch')


def show_sheet_preview(sheet_name, df):
    st.header(sheet_name)
    if df.empty:
        st.warning('This sheet contains no usable data after cleaning.')
        return

    st.subheader('Sample data')
    st.dataframe(df.head(20), width='stretch')

    date_col = parse_dates(df)
    numeric_columns = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]

    if date_col and numeric_columns:
        selected = st.selectbox('Choose a numeric series to plot', numeric_columns, index=0)
        show_chart(df, date_col, selected, title=f'{sheet_name}: {selected} over time')
    elif len(numeric_columns) > 0:
        selected = st.multiselect('Choose numeric columns to plot', numeric_columns, default=numeric_columns[:3])
        if selected:
            show_chart(df, df.columns[0], selected, title=f'{sheet_name}: selected series', chart_type='bar')
    else:
        st.info('No numeric data available for charting in this sheet.')

    if st.checkbox('Show descriptive statistics', key=f'stats_{sheet_name}'):
        st.write(df.describe(include='all'))


def show_dashboard(sheet_data):
    st.title('KNBS Data Dashboard')
    st.write('Interactive exploration of KNBS macro, price, trade, and financial indicators from the provided workbook.')

    highlights = [
        ('Inflation', 'Inflation', ['KENYA CPI', 'NAIROBI COMBINED']),
        ('Fuel Prices', 'Fuel Prices', ['Motor Gasoline Premium (KSh per Litre)', 'Light Diesel Oil (KSh per Litre)']),
        ('Interest Rates', 'Interest Rates', ['Central bank Rates', 'Average Yield Rates 91-Days Treasury Bills']),
        ('Forex Reserves', 'Forex Reserves', ['NET FOREIGN EXCHANGE RESERVES', 'GROSS TOTAL']),
        ('Mobile Money', 'Mobile Money Trans', ['Value (KSh billions)', 'QUASI -MONEY']),
        ('Trade', 'External Trade', ['Total Exports', 'Total Imports']),
    ]

    for section, sheet_name, candidates in highlights:
        if sheet_name in sheet_data:
            df = sheet_data[sheet_name]
            good_col = find_best_numeric(df, candidates)
            date_col = parse_dates(df)
            if good_col and date_col:
                with st.container():
                    st.subheader(f'{section}')
                    show_chart(df, date_col, good_col, title=f'{section} trend')

    st.markdown('---')
    st.subheader('Sheet overview')
    summary = []
    for sheet_name, df in sheet_data.items():
        summary.append((sheet_name, len(df), ', '.join(str(c) for c in df.columns[:4])))

    overview = pd.DataFrame(summary, columns=['Sheet', 'Rows', 'First columns'])
    st.dataframe(overview, width='stretch')


def show_inflation_page(sheet_data):
    st.title('Inflation & Consumer Prices')
    st.write('This page focuses on Kenya CPI and inflation trends from the KNBS workbook.')

    if 'Inflation' in sheet_data:
        df = sheet_data['Inflation']
        date_col = parse_dates(df)
        if date_col:
            candidates = ['KENYA CPI', 'NAIROBI COMBINED']
            available = [c for c in candidates if c in df.columns and pd.api.types.is_numeric_dtype(df[c])]
            if available:
                show_chart(df, date_col, available[0], title='Kenya CPI inflation trend')
                if len(available) > 1:
                    show_chart(df, date_col, available[1], title='Nairobi combined inflation trend')
    if 'CPI' in sheet_data:
        st.markdown('### CPI indices')
        df = sheet_data['CPI']
        date_col = parse_dates(df)
        if date_col:
            available = [c for c in ['KENYA CPI', 'NAIROBI COMBINED', 'NAIROBI UPPER INCOME'] if c in df.columns]
            if available:
                show_chart(df, date_col, available[0], title='CPI overall trend')
    if 'CPI Old' in sheet_data:
        st.markdown('### Historical CPI and underlying inflation')
        df = sheet_data['CPI Old']
        date_col = parse_dates(df)
        if date_col and 'OVERALL KENYA INDEX' in df.columns:
            show_chart(df, date_col, 'OVERALL KENYA INDEX', title='Historical Kenya CPI')


def show_fuel_page(sheet_data):
    st.title('Fuel Prices')
    st.write('Track fuel price movements for gasoline, diesel, kerosene, and LPG.')
    if 'Fuel Prices' in sheet_data:
        df = sheet_data['Fuel Prices']
        date_col = parse_dates(df)
        if date_col:
            fuel_cols = [
                'Motor Gasoline Premium (KSh per Litre)',
                'Light Diesel Oil (KSh per Litre)',
                'I lluminating Kerosene (KSh per Litre)',
                'L.P.G (KSh per 13 kg)'
            ]
            available = [c for c in fuel_cols if c in df.columns]
            if available:
                show_chart(df, date_col, available[0], title=available[0])
                if len(available) > 1:
                    show_chart(df, date_col, available[1], title=available[1])
        st.subheader('Fuel price data sample')
        st.dataframe(df.head(15), width='stretch')
    else:
        st.warning('Fuel Prices sheet is not available in the workbook.')


def show_trade_page(sheet_data):
    st.title('Trade and Foreign Exchange')
    st.write('Explore exports, imports and reserves from KNBS trade data.')
    if 'External Trade' in sheet_data:
        df = sheet_data['External Trade']
        date_col = parse_dates(df)
        if date_col:
            if 'Total Exports' in df.columns and 'Total Imports' in df.columns:
                show_chart(df, date_col, ['Total Exports', 'Total Imports'], title='Trade exports vs imports')
            elif 'Total Exports' in df.columns:
                show_chart(df, date_col, 'Total Exports', title='Total Exports')
        st.subheader('External Trade sample')
        st.dataframe(df.head(15), width='stretch')
    if 'Forex Reserves' in sheet_data:
        st.markdown('### Foreign exchange reserves')
        df = sheet_data['Forex Reserves']
        date_col = parse_dates(df)
        if date_col and 'NET FOREIGN EXCHANGE RESERVES' in df.columns:
            show_chart(df, date_col, 'NET FOREIGN EXCHANGE RESERVES', title='Net forex reserves')
        st.dataframe(df.head(15), width='stretch')


def show_mobile_money_page(sheet_data):
    st.title('Digital Finance & Mobile Money')
    st.write('Mobile money transaction levels and values from the KNBS workbook.')
    if 'Mobile Money Trans' in sheet_data:
        df = sheet_data['Mobile Money Trans']
        date_col = parse_dates(df)
        if date_col:
            if 'Value (KSh billions)' in df.columns:
                show_chart(df, date_col, 'Value (KSh billions)', title='Mobile money value (KSh billions)')
        st.subheader('Mobile Money sample')
        st.dataframe(df.head(15), width='stretch')
    else:
        st.warning('Mobile Money Trans sheet is not available in the workbook.')


def show_logistics_page(sheet_data):
    st.title('Logistics & Infrastructure')
    st.write('Transport, ports, logistics and infrastructure indicators from KNBS.')

    if 'SGR' in sheet_data:
        sgr_df = sheet_data['SGR']
        date_col = parse_dates(sgr_df)
        if date_col:
            if 'Tonnage' in sgr_df.columns:
                show_chart(sgr_df, date_col, 'Tonnage', title='SGR freight throughput')
            if 'Passengers' in sgr_df.columns:
                show_chart(sgr_df, date_col, 'Passengers', title='SGR passenger volumes')
        st.subheader('SGR data sample')
        st.dataframe(sgr_df.head(15), width='stretch')
    else:
        st.warning('SGR sheet is not available in the workbook.')

    if 'Port' in sheet_data:
        port_df = sheet_data['Port']
        date_col = parse_dates(port_df)
        if date_col:
            port_cols = [c for c in ['Cargo Volume', 'Container Throughput', 'Ship Calls'] if c in port_df.columns]
            if port_cols:
                for col in port_cols[:2]:
                    show_chart(port_df, date_col, col, title=f'Port {col} trend')
        st.subheader('Port data sample')
        st.dataframe(port_df.head(15), width='stretch')

    if 'Vehicle Reg' in sheet_data:
        vh_df = sheet_data['Vehicle Reg']
        date_col = parse_dates(vh_df)
        if date_col and 'Total Registrations' in vh_df.columns:
            show_chart(vh_df, date_col, 'Total Registrations', title='Vehicle registrations')
            st.subheader('Vehicle registration data sample')
            st.dataframe(vh_df.head(15), width='stretch')


def show_sector_summary_page(sheet_data):
    st.title('Sector Summary')
    st.write('Cross-sector summary of key KNBS indicators for macro, trade, finance, energy, and logistics.')

    latest = {}

    if 'Inflation' in sheet_data:
        df = sheet_data['Inflation']
        date_col = parse_dates(df)
        if date_col and 'KENYA CPI' in df.columns:
            value = df.loc[df[date_col].idxmax(), 'KENYA CPI']
            latest['Kenya CPI'] = f"{value:.2f}" if pd.notna(value) else 'N/A'

    if 'Fuel Prices' in sheet_data:
        df = sheet_data['Fuel Prices']
        date_col = parse_dates(df)
        if date_col and 'Motor Gasoline Premium (KSh per Litre)' in df.columns:
            value = df.loc[df[date_col].idxmax(), 'Motor Gasoline Premium (KSh per Litre)']
            latest['Gasoline Price'] = f"{value:.2f} KSh/L" if pd.notna(value) else 'N/A'

    if 'Forex Reserves' in sheet_data:
        df = sheet_data['Forex Reserves']
        date_col = parse_dates(df)
        if date_col and 'NET FOREIGN EXCHANGE RESERVES' in df.columns:
            value = df.loc[df[date_col].idxmax(), 'NET FOREIGN EXCHANGE RESERVES']
            latest['Forex Reserves'] = f"{value:,.0f} KSh" if pd.notna(value) else 'N/A'

    if 'Mobile Money Trans' in sheet_data:
        df = sheet_data['Mobile Money Trans']
        date_col = parse_dates(df)
        if date_col and 'Value (KSh billions)' in df.columns:
            value = df.loc[df[date_col].idxmax(), 'Value (KSh billions)']
            latest['Mobile Money Value'] = f"{value:.2f} B KSh" if pd.notna(value) else 'N/A'

    if 'External Trade' in sheet_data:
        df = sheet_data['External Trade']
        date_col = parse_dates(df)
        if date_col:
            if 'Total Exports' in df.columns and 'Total Imports' in df.columns:
                exports = df.loc[df[date_col].idxmax(), 'Total Exports']
                imports = df.loc[df[date_col].idxmax(), 'Total Imports']
                if pd.notna(exports) and pd.notna(imports):
                    latest['Trade Balance'] = f"{(exports - imports):,.0f}"
            elif 'Total Exports' in df.columns:
                exports = df.loc[df[date_col].idxmax(), 'Total Exports']
                latest['Total Exports'] = f"{exports:,.0f}"

    if 'SGR' in sheet_data:
        sgr_df = sheet_data['SGR']
        date_col = parse_dates(sgr_df)
        if date_col and 'Tonnage' in sgr_df.columns:
            value = sgr_df.loc[sgr_df[date_col].idxmax(), 'Tonnage']
            latest['SGR Tonnage'] = f"{value:,.0f}" if pd.notna(value) else 'N/A'

    if 'Vehicle Reg' in sheet_data:
        vh_df = sheet_data['Vehicle Reg']
        date_col = parse_dates(vh_df)
        if date_col and 'Total Registrations' in vh_df.columns:
            value = vh_df.loc[vh_df[date_col].idxmax(), 'Total Registrations']
            latest['Vehicle Registrations'] = f"{value:,.0f}" if pd.notna(value) else 'N/A'

    if latest:
        cols = st.columns(min(len(latest), 4))
        for col, (label, value) in zip(cols, latest.items()):
            col.metric(label, value)
    else:
        st.info('No sector summary metrics available from the workbook.')

    st.markdown('---')
    st.subheader('Recent sector indicators')
    for sheet_name in ['Inflation', 'Fuel Prices', 'External Trade', 'Forex Reserves', 'Mobile Money Trans', 'SGR', 'Vehicle Reg']:
        if sheet_name in sheet_data:
            df = sheet_data[sheet_name]
            date_col = parse_dates(df)
            numeric = find_best_numeric(df)
            if date_col and numeric:
                show_chart(df, date_col, numeric, title=f'{sheet_name}: {numeric}')
                st.write(f'*{sheet_name} sample*')
                st.dataframe(df[[date_col, numeric]].tail(8), width='stretch')

    latest_values = {}

    if 'Inflation' in sheet_data:
        df = sheet_data['Inflation']
        date_col = parse_dates(df)
        if date_col and 'KENYA CPI' in df.columns:
            latest = df.loc[df[date_col].idxmax(), 'KENYA CPI']
            latest_values['Latest Kenya CPI'] = f"{latest:.2f}" if pd.notna(latest) else 'N/A'

    if 'Forex Reserves' in sheet_data:
        df = sheet_data['Forex Reserves']
        date_col = parse_dates(df)
        if date_col and 'NET FOREIGN EXCHANGE RESERVES' in df.columns:
            latest = df.loc[df[date_col].idxmax(), 'NET FOREIGN EXCHANGE RESERVES']
            latest_values['Net Forex Reserves'] = f"{latest:,.0f}" if pd.notna(latest) else 'N/A'

    if 'Mobile Money Trans' in sheet_data:
        df = sheet_data['Mobile Money Trans']
        date_col = parse_dates(df)
        if date_col and 'Value (KSh billions)' in df.columns:
            latest = df.loc[df[date_col].idxmax(), 'Value (KSh billions)']
            latest_values['Mobile Money Value'] = f"{latest:.2f} B KSh" if pd.notna(latest) else 'N/A'

    if 'Interest Rates' in sheet_data:
        df = sheet_data['Interest Rates']
        date_col = parse_dates(df)
        if date_col and 'Central bank Rates' in df.columns:
            latest = df.loc[df[date_col].idxmax(), 'Central bank Rates']
            latest_values['Central Bank Rate'] = f"{latest:.2f}%" if pd.notna(latest) else 'N/A'

    if 'External Trade' in sheet_data:
        df = sheet_data['External Trade']
        date_col = parse_dates(df)
        if date_col and 'Total Exports' in df.columns:
            latest = df.loc[df[date_col].idxmax(), 'Total Exports']
            latest_values['Total Exports'] = f"{latest:,.0f}" if pd.notna(latest) else 'N/A'

    if latest_values:
        cols = st.columns(len(latest_values))
        for col, (label, value) in zip(cols, latest_values.items()):
            col.metric(label, value)
    else:
        st.info('No summary metrics available for the current workbook.')

    st.markdown('---')

    st.subheader('Trend highlights')
    if 'Inflation' in sheet_data:
        df = sheet_data['Inflation']
        date_col = parse_dates(df)
        if date_col and 'KENYA CPI' in df.columns:
            show_chart(df, date_col, 'KENYA CPI', title='Kenya CPI trend')
    if 'Forex Reserves' in sheet_data:
        df = sheet_data['Forex Reserves']
        date_col = parse_dates(df)
        if date_col and 'NET FOREIGN EXCHANGE RESERVES' in df.columns:
            show_chart(df, date_col, 'NET FOREIGN EXCHANGE RESERVES', title='Foreign exchange reserves trend')


sheet_data = get_data()
page = st.sidebar.radio('Navigation', ['Dashboard', 'Macro Overview', 'Sector Summary', 'Inflation', 'Fuel Prices', 'Trade', 'Mobile Money', 'Logistics & Infrastructure', 'Sheet Explorer', 'About'])

if page == 'Dashboard':
    show_dashboard(sheet_data)

elif page == 'Macro Overview':
    show_macro_overview_page(sheet_data)

elif page == 'Sector Summary':
    show_sector_summary_page(sheet_data)

elif page == 'Inflation':
    show_inflation_page(sheet_data)

elif page == 'Fuel Prices':
    show_fuel_page(sheet_data)

elif page == 'Trade':
    show_trade_page(sheet_data)

elif page == 'Mobile Money':
    show_mobile_money_page(sheet_data)

elif page == 'Logistics & Infrastructure':
    show_logistics_page(sheet_data)

elif page == 'Sheet Explorer':
    st.title('KNBS Sheet Explorer')
    sheets = list(sheet_data.keys())
    if not sheets:
        st.error('Could not find KNBS Data.xlsx in the same folder as app.py.')
    else:
        selected_sheet = st.selectbox('Choose a sheet', sheets)
        show_sheet_preview(selected_sheet, sheet_data[selected_sheet])

else:
    st.title('About this Website')
    st.write('This Streamlit website reads the `KNBS Data.xlsx` workbook and exposes every sheet as interactive data and charts.')
    st.write('Use the sidebar to switch between the dashboard and sheet explorer.')
    st.write('Data is loaded directly from the workbook, so any updates to the Excel file are reflected when the app is refreshed.')
    st.write('The site now includes dedicated pages for Macro Overview, Sector Summary, Inflation, Fuel Prices, Trade, Mobile Money, and Logistics & Infrastructure.')
