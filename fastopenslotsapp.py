import sys

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import pytz
from datetime import datetime, timedelta
import calendar
from st_aggrid import GridOptionsBuilder, AgGrid, JsCode
from st_aggrid.shared import GridUpdateMode, DataReturnMode, ColumnsAutoSizeMode, AgGridTheme, ExcelExportMode
from st_aggrid.AgGridReturn import AgGridReturn
import json
import os
import numpy as np

@st.cache_data
def load_excel(file_path, usecols=None, **kwargs):
    """ Load an Excel file if it exists. Optionally specify columns to load. """
    print(f"Checking for file: {file_path}")
    if os.path.exists(file_path):
        print(f"Loading {file_path}")
        return pd.read_excel(file_path, usecols=usecols, **kwargs)
    else:
        print(f"File {file_path} not found.")
        return None

def find_last_working_day(start_date):
    """ Helper function to find the last working day before a given start date. """
    current_date = start_date
    while True:
        # Assuming Monday to Friday are working days (customize as needed)
        if current_date.weekday() < 5:  # Monday is 0 and Friday is 4
            return current_date
        current_date -= timedelta(days=1)

st.set_page_config(layout="wide")

# Apply custom CSS to adjust the sidebar and main content width
st.markdown(
    """
    <style>
    /* Reduce the width of the sidebar */
    [data-testid="stSidebar"] {
        width: 100px; 
        background-color: #cc0641; 
    }

    /* Make the main content fill more of the screen */
    .css-1lcbmhc {
        max-width: calc(100% - 200px);  
        margin-left: -200px;
    }

    /* Sidebar filter headers */
    [data-testid="stSidebar"] .stSelectbox > label {
        color: white;
        font-weight: bold;
    }

    /* Sidebar statistics styles */
    .sidebar-stats-box {
        color: white;  /* Text color */
        font-weight: bold;  /* Bold text */
        border: 2px solid white ;  /* Add a border around the box */
        background-color: #cc0641;  /* Set background to transparent or keep it to match the sidebar */
        padding: 10px;  /* Padding inside the box */
        margin-bottom: 5px;  /* Space below each box */
        border-radius: 5px;  /* Optional: Rounded corners */
    }

    .custom-title {
        font-size: 2em;  /* Adjust font size */
        color: #cc0641;  /* Change text color to match your theme */
        font-weight: bold;  /* Make text bold */
        text-align: center;  /* Center align the title */
        margin-top: -50px;  /* Move the title higher by using a negative margin */
        margin-bottom: 20px;  /* Add space below the title */
        background-color: #f0f2f6;  /* Optional: Add a subtle background color */
        padding: 10px;  /* Add padding around the title */
        border-radius: 10px;  /* Optional: Rounded corners for the background */
    }
    
    /* Circle styles for color-coded labels */
    .circle {
        height: 15px;
        width: 15px;
        display: inline-block;
        border-radius: 50%;
        margin-right: 10px;
    }
    .red-circle {
        background-color: #cc0641;
    }
    .orange-circle {
        background-color: #f1b84b;
    }
    .green-circle {
        background-color: #95cd41;
    }
    
    .label-container {
        text-align: center;  /* Center align the labels */
        margin-bottom: 20px;  /* Add space below the labels */
    }
    
    .label-text {
        display: inline-block;
        vertical-align: middle;
        font-size: 1em;  /* Adjust font size */
        margin-right: 20px;  /* Space between labels */
    }

    </style>
    """,
    unsafe_allow_html=True
)
# Streamlit App
st.markdown('<h1 class="custom-title">SLOT AVAILABILITY REPORT</h1>', unsafe_allow_html=True)
# Define the time zones
st.markdown(
    """
    <div class="label-container">
        <span class="label-text"><span class="circle red-circle"></span>La agenda está mal configurada</span>
        <span class="label-text"><span class="circle orange-circle"></span>La agenda está bien configurada, pero llena</span>
        <span class="label-text"><span class="circle green-circle"></span>La agenda está bien configurada y disponible</span>
    </div>
    """,
    unsafe_allow_html=True
)
def get_first_iso_week_start_date_current_month():
    # Get the current date
    today = datetime.today()
    
    # Get the first day of the current month with the time set to 00:00:00
    first_day_of_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    
    # Check if the first day of the month is a Sunday
    if first_day_of_month.weekday() == 6:  # Sunday is represented by 6 in weekday()
        # If Sunday, move to the next Monday
        first_day_of_month += timedelta(days=1)
    
    # Get the ISO calendar week and weekday of the first day of the month
    iso_year, iso_week, iso_weekday = first_day_of_month.isocalendar()
    
    # Calculate the difference to get back to the Monday of that ISO week
    # ISO weeks start on Monday (iso_weekday = 1), so subtract the days to go back to Monday
    start_date = first_day_of_month - timedelta(days=iso_weekday - 1)
    
    # Ensure the time part is set to 00:00:00
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)
    
    return start_date

def get_last_iso_week_end_date_current_month():
    # Get the current date
    today = datetime.today()
    
    # Get the last day of the current month with the time set to 00:00:00
    last_day_of_month = today.replace(day=calendar.monthrange(today.year, today.month)[1], hour=0, minute=0, second=0, microsecond=0)
    
    # Get the ISO calendar week and weekday of the last day of the month
    iso_year, iso_week, iso_weekday = last_day_of_month.isocalendar()
    
    # Calculate the difference to get to Sunday of that ISO week
    # ISO weeks end on Sunday (iso_weekday = 7), so add the days to go to Sunday
    end_date = last_day_of_month + timedelta(days=(7 - iso_weekday))
    
    # Ensure the time part is set to 00:00:00
    end_date = end_date.replace(hour=0, minute=0, second=0, microsecond=0)
    
    return end_date

folder_path = 'shiftslots'

# Example usage: dynamically calculate start and end dates for the current month
start_date = get_first_iso_week_start_date_current_month()
end_date = get_last_iso_week_end_date_current_month()

start_iso_year, start_iso_week, _ = start_date.isocalendar()
end_iso_year, end_iso_week, _ = end_date.isocalendar()
current_iso_year, current_iso_week, _ = datetime.now().isocalendar()

def get_start_and_end_of_current_month():
    # Get the current date
    today = datetime.today()
    
    # Start date is the 1st day of the current month
    month_start_date = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    
    # End date is the last day of the current month
    last_day_of_month = calendar.monthrange(today.year, today.month)[1]  # Get last day of current month
    month_end_date = today.replace(day=last_day_of_month, hour=23, minute=59, second=59, microsecond=999999)
    
    return month_start_date, month_end_date

# Example usage to get the start and end dates of the current month
month_start_date, month_end_date = get_start_and_end_of_current_month()
current_date = datetime.now()
yesterday_date = current_date - timedelta(days=1)

# Determine if today is Monday
if current_date.weekday() == 0:  # Monday
    print("Today is Monday, directly finding last Friday's file...")
    # On Monday, skip Sunday and Saturday
    last_working_day = find_last_working_day(current_date - timedelta(days=3))
else:
    # On other days, start with today
    last_working_day = find_last_working_day(current_date)

today_file_name = f"shiftslots_{current_date.strftime('%Y-%m-%d')}.xlsx"
yesterday_file_name = f"shiftslots_{yesterday_date.strftime('%Y-%m-%d')}.xlsx"
sep6_file_name = f"shiftslots_{month_start_date.strftime('%Y-%m-%d')}.xlsx"
last_working_day_file_name = f"shiftslots_{last_working_day.strftime('%Y-%m-%d')}.xlsx"

shift_slots = load_excel(os.path.join(folder_path, today_file_name))

# Fallback to yesterday's file if today’s file is missing and it’s not Monday
if shift_slots is None and current_date.weekday() != 0:
    print("Today's file not found, trying yesterday's file...")
    shift_slots = load_excel(os.path.join(folder_path, yesterday_file_name))

# Fallback to last working day’s file if both today’s and yesterday’s files are missing
if shift_slots is None:
    print("No file found for today or yesterday, trying the last working day's file...")
    shift_slots = load_excel(os.path.join(folder_path, last_working_day_file_name))

# Stop if no file is found after all attempts
if shift_slots is None:
    st.error("No file found for today, yesterday, or the last working day.")
    st.stop()
shift_slots_yesterday = load_excel(os.path.join(folder_path, yesterday_file_name))
if shift_slots_yesterday is None:
    print("Yesterday's file not found, finding the last working day...")
    last_working_day_yesterday = find_last_working_day(yesterday_date)
    last_working_day_yesterday_file_name = f"shiftslots_{last_working_day_yesterday.strftime('%Y-%m-%d')}.xlsx"
    shift_slots_yesterday = load_excel(os.path.join(folder_path, last_working_day_yesterday_file_name))

shift_slots_sep6 = load_excel(os.path.join(folder_path, sep6_file_name))
hcp_shift_slots = load_excel('output/hcpshiftslots.xlsx')
hcm = load_excel('output/hcm_sf_merged.xlsx')
clock= load_excel('output/clock.xlsx')
# Assuming `shift_slots['iso_week']` is a list of ISO weeks
available_weeks = sorted(shift_slots['iso_week'].unique())
# Find the index of the current ISO week in the list
if current_iso_week in available_weeks:
    current_week_index = available_weeks.index(current_iso_week)
else:
    current_week_index = 0  # Fallback to the first week if current week is not available

# Sidebar filters (all converted to single-selection using selectbox)
iso_week_filter = st.sidebar.selectbox('Select ISO Week', available_weeks, index=current_week_index)

# Calculate the previous ISO week and year based on the selected ISO week
selected_iso_year = datetime.now().year  # Assuming current year, adjust if you have a different dataset
previous_iso_week = iso_week_filter - 1
previous_iso_year = selected_iso_year
if previous_iso_week == 0:
    previous_iso_year -= 1
    previous_iso_week = 52 if (pd.Timestamp(f"{previous_iso_year}-12-28").isocalendar()[1] == 52) else 53

# Sidebar filter for Region
region_list = sorted(shift_slots['Region'].dropna().unique().tolist())
region_options = ["All"] + region_list

selected_region = st.sidebar.selectbox(
    'Select Region:',
    options=region_options,
    index=0,  # Default to "All"
    help="Select a region or 'All' to view data for all regions."
)
# Filter data based on the selected region
if selected_region == "All":
    filtered_shift_slots_by_region = shift_slots
else:
    filtered_shift_slots_by_region = shift_slots[shift_slots['Region'] == selected_region]

# Sidebar filter for Area based on filtered data by Region
area_list = sorted(filtered_shift_slots_by_region['Area'].dropna().unique().tolist())
area_options = ["All"] + area_list

selected_area = st.sidebar.selectbox(
    'Select Area:',
    options=area_options,
    index=0,  # Default to "All"
    help="Select an area or 'All' to view data for all areas."
)

# Further filter data based on the selected area
if selected_area == "All":
    filtered_shift_slots_by_area = filtered_shift_slots_by_region
else:
    filtered_shift_slots_by_area = filtered_shift_slots_by_region[filtered_shift_slots_by_region['Area'] == selected_area]

# Sidebar filter for Shop based on filtered data by Region and Area
shop_list = sorted(filtered_shift_slots_by_area['Shop[Name]'].dropna().unique().tolist())
shop_options = ["All"] + shop_list

selected_shop = st.sidebar.selectbox(
    'Select Shop:',
    options=shop_options,
    index=0,  # Default to "All"
    help="Select a shop or 'All' to view data for all shops."
)

# Global Filter Function for other datasets (with iso_week filter)
@st.cache_data
def filter_data(data, iso_week_filter, selected_region, selected_area, selected_shop, week_column='iso_week'):
    filtered_data = data[data[week_column] == iso_week_filter]

    if selected_region != "All":
        filtered_data = filtered_data[filtered_data['Region'] == selected_region]

    if selected_area != "All":
        filtered_data = filtered_data[filtered_data['Area'] == selected_area]

    if selected_shop != "All":
        filtered_data = filtered_data[filtered_data['Shop[Name]'] == selected_shop]

    return filtered_data

# Special filter function for HCM data (no iso_week filter)
@st.cache_data
def filter_hcm_data(data, selected_region, selected_area, selected_shop):
    filtered_data = data.copy()  # Copy data without applying iso_week filter

    if selected_region != "All":
        filtered_data = filtered_data[filtered_data['Region'] == selected_region]

    if selected_area != "All":
        filtered_data = filtered_data[filtered_data['Area'] == selected_area]

    if selected_shop != "All":
        filtered_data = filtered_data[filtered_data['Shop Name'] == selected_shop]

    return filtered_data

@st.cache_data
def filter_hcp_shift_slots(data, selected_region, selected_area, selected_shop):
    filtered_data = data.copy()  # Copy data without applying iso_week filter

    if selected_region != "All":
        filtered_data = filtered_data[filtered_data['Region'] == selected_region]

    if selected_area != "All":
        filtered_data = filtered_data[filtered_data['Area'] == selected_area]

    if selected_shop != "All":
        filtered_data = filtered_data[filtered_data['Shop[Name]'] == selected_shop]

    return filtered_data

# Apply the filters fo r the other datasets
filtered_data = filter_data(shift_slots, iso_week_filter, selected_region, selected_area, selected_shop, 'iso_week')
filtered_hcp_shift_slots = filter_hcp_shift_slots(hcp_shift_slots, selected_region, selected_area, selected_shop)
weekly_shift_slots = filter_hcp_shift_slots(shift_slots, selected_region, selected_area, selected_shop)
weekly_shift_slots_yesterday = filter_hcp_shift_slots(shift_slots_yesterday, selected_region, selected_area, selected_shop)
weekly_shift_sep6 = filter_hcp_shift_slots(shift_slots_sep6, selected_region, selected_area, selected_shop)

# Apply the filters to HCM data (without iso_week filter)
filtered_hcm = filter_hcm_data(hcm, selected_region, selected_area, selected_shop)
filtered_clock = filter_data(clock,iso_week_filter, selected_region, selected_area, selected_shop, 'iso_week')
filtered_clock_noiso = filter_hcp_shift_slots(clock, selected_region, selected_area, selected_shop)

# Filter data for the previous ISO week without applying the current filters (directly from shift_slots)
previous_week_data = shift_slots[
    (shift_slots['iso_week'] == previous_iso_week) &
    (shift_slots['Region'] == selected_region if selected_region != "All" else True) &
    (shift_slots['Area'] == selected_area if selected_area != "All" else True) &
    (shift_slots['Shop[Name]'] == selected_shop if selected_shop != "All" else True)
]

# Calculate Open Hours for the current and previous weeks
open_hours_this_week = filtered_data['OpenHours'].sum()
open_hours_last_week = previous_week_data['OpenHours'].sum()

# Calculate percentage change from last week, with checks to prevent division by zero
if open_hours_last_week != 0:
    change_from_last_week = ((open_hours_this_week - open_hours_last_week) / open_hours_last_week) * 100
else:
    change_from_last_week = 0  # Or handle differently, depending on your needs

# Calculate the start and end dates for the selected ISO week
selected_week_start = pd.Timestamp(selected_iso_year, 1, 1) + pd.offsets.Week(weekday=0) * (int(iso_week_filter) - 1)
selected_week_end = selected_week_start + pd.offsets.Week(weekday=6)
today = pd.Timestamp(datetime.now().date())
end_of_month = today.replace(day=1) + pd.offsets.MonthEnd(0)
# Calculate "Open Hours for the month to go" using the entire dataset
month_to_go_data = shift_slots[(shift_slots['date'] >= today) & (shift_slots['date'] <= end_of_month)]
open_hours_month_to_go = month_to_go_data['OpenHours'].sum()


# Determine the best configured region      

best_configured_region = shift_slots.groupby('Region')['SaturationPercentage'].mean().idxmax() if not filtered_data.empty else 'N/A'
st.sidebar.markdown(f"<div class='sidebar-stats-box'>Open hours for the selected week: {open_hours_this_week:,.0f}</div>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div class='sidebar-stats-box'>Change from last week: {change_from_last_week:,.0f}%</div>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div class='sidebar-stats-box'>Open hours for month to go: {open_hours_month_to_go:,.0f}</div>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div class='sidebar-stats-box'>Best configured region: {best_configured_region}</div>", unsafe_allow_html=True)
# Ensure that only numeric columns are included for aggregation
numeric_cols = ['OpenHours', 'TotalHours', 'BlockedHoursPercentage']
filtered_data_numeric = filtered_data[numeric_cols + ['day', 'weekday', 'GT_ShopCode__c', 'Shop[Name]']]

# Aggregating data by GT_ShopCode__c, Shop[Name], date, and weekday
aggregated_data = filtered_data.groupby(['GT_ShopCode__c', 'Shop[Name]', 'date', 'weekday']).agg(
    OpenHours=('OpenHours', 'sum'),
    TotalHours=('TotalHours', 'sum'),
    BlockedHoursPercentage=('BlockedHoursPercentage', 'mean')
).reset_index()

aggregated_data['date'] = pd.to_datetime(aggregated_data['date']).dt.date
check= aggregated_data[(aggregated_data['GT_ShopCode__c'] == '240')]

# Aggregating data by GT_ShopCode__c, Shop[Name], date, and weekday
hcp_data = filtered_hcp_shift_slots.groupby(['PersonalNumberKey', 'GT_ServiceResource__r.Name', 'GT_ShopCode__c', 'Shop[Name]', 'ShiftDate', 'weekday']).agg(
    AvailableHours=('ShiftDurationHoursAdjusted', 'sum'),
    BlockedHours=('AbsenceDurationHours', 'sum'),
).reset_index()

tab6, tab1, tab2,tab4, tab3,tab5 = st.tabs(["Weekly Change Analysis", "Open Hours / Total Hours", "Blocked Hours %", "HCM vs SF", "ACT vs SF","REX"])

with tab1: 
    if aggregated_data.empty:
        st.warning("No shops found for the selected filter criteria.")

    st.markdown(''':green[ *****Total Hours***: Todos los turnos configurados*]''')
    st.markdown(''':green[ *****Open Hours***: Turnos abiertos luego de descontar horas bloqueadas y horas de cita*]''')

    # Adjust the pivot table to exclude GT_ShopCode__c and SaturationPercentage
    pivot_table = aggregated_data.pivot_table(
        index=['Shop[Name]'],
        columns=['date', 'weekday'],
        values=['OpenHours', 'TotalHours'],
        aggfunc='sum',
        fill_value=0  
    )

    # Flatten the columns
    pivot_table.columns = [f"{col[0]}_{col[1]}_{col[2]}" for col in pivot_table.columns.to_flat_index()]
    pivot_table_reset = pivot_table.reset_index()

    # Format all numeric columns to one decimal point
    numeric_columns_in_pivot = [col for col in pivot_table_reset.columns if any(nc in col for nc in ['OpenHours', 'TotalHours'])]
    pivot_table_reset[numeric_columns_in_pivot] = pivot_table_reset[numeric_columns_in_pivot].round(1)

    # Create the DataFrame (df)
    df = pivot_table_reset

    # Ensure no spaces in field names in df, replacing spaces with underscores or removing them
    df.columns = [col.replace(' ', '_') for col in df.columns]
    js_code = JsCode("""
    function(params) {
        var totalHoursField = params.colDef.field.replace('OpenHours', 'TotalHours');
        var openHoursField = params.colDef.field.replace('TotalHours', 'OpenHours');
        var totalHoursValue = params.data[totalHoursField];
        var openHoursValue = params.data[openHoursField];

        if (params.data['Shop[Name]'] === 'Total') {
            return {'font-weight': 'bold', 'backgroundColor': '#e0e0e0'};  // Make the total row bold and set a light background color for visibility
        } else if (totalHoursValue === 0) {
            return {'backgroundColor': '#cc0641', 'color': 'white'};  // Red background and white text for TotalHours = 0
        } else if (openHoursValue !== 0 && totalHoursValue !== 0) {
            return {'backgroundColor': '#95cd41'};  // Green for OpenHours != 0 and TotalHours != 0
        } else if (openHoursValue === 0 && totalHoursValue !== 0) {
            return {'backgroundColor': '#f1b84b'};  // Orange for OpenHours = 0 and TotalHours != 0
        } else {
            return null; 
        }
    }
    """)

    custom_css = {
        ".ag-header-cell": {
            "background-color": "#cc0641 !important",  # Ensure entire cell background changes
            "color": "white !important",
            "font-weight": "bold",
            "padding": "4px"  # Reduce padding to make headers more compact
        },
        ".ag-header-group-cell": {  # Style for merged/group headers
            "background-color": "#cc0641 !important",
            "color": "white !important",
            "font-weight": "bold",
        },
        ".ag-cell": {
            "padding": "2px",  # Reduce padding inside cells to make them more compact
            "font-size": "12px"  # Reduce font size for a more compact look
        },
        ".ag-header": {
            "height": "35px",  # Reduce header height
        },
        ".ag-theme-streamlit .ag-row": {
            "max-height": "30px"  # Adjust max height for rows to be more compact
        },
        ".ag-theme-streamlit .ag-menu-option-text, .ag-theme-streamlit .ag-filter-body-wrapper, .ag-theme-streamlit .ag-input-wrapper, .ag-theme-streamlit .ag-icon": {
            "font-size": "6px !important"  # Reduce font size for filter options and ensure it's applied
        },
        ".ag-theme-streamlit .ag-root-wrapper": {
            "border": "2px solid #cc0641",  # Add outer border with specified color
            "border-radius": "5px"  # Optional: Rounded corners for the outer border
        }
    }

    # Example column definition with flex and resizable properties
    columnDefs = [
        {
            "headerName": "Shop Name",
            "field": "Shop[Name]", 
            "resizable": True,
            "flex": 2,  # Adjust flex value to make this column wider
            "minWidth": 150,  # Set a minimum width for columns
            "filter": 'agTextColumnFilter',  # Set filter type to text for shop name
        },
    ]

    # Append dynamic column definitions with conditional formatting for OpenHours and TotalHours
    for column in df.columns[1:]:  # Start from 1 to skip Shop_Name
        if 'OpenHours' in column:
            headerName = column.split('_')[1] + ' (' + column.split('_')[2] + ')'
            columnDefs.append({
                "headerName": headerName,
                "children": [
                    {
                        "field": column,
                        "headerName": "Open Hours",
                        "valueFormatter": "x.toFixed(1)",
                        "resizable": True,
                        "flex": 1,
                        "cellStyle": js_code                },
                    {
                        "field": column.replace('OpenHours', 'TotalHours'),
                        "headerName": "Total Hours",
                        "valueFormatter": "x.toFixed(1)",
                        "resizable": True,
                        "flex": 1,   
                        "cellStyle": js_code                  }
                ]
            })

    # Calculate totals for numeric columns
    total_row = {
        'Shop[Name]': 'Total'
    }

    # Iterate over the numeric columns to compute totals
    for col in numeric_columns_in_pivot:
        total_row[col] =  f"{int(df[col].sum().round(0)):,}"


    # Convert total_row to DataFrame
    total_df = pd.DataFrame(total_row, index=[0])

    df_with_totals = pd.concat([total_df, df], ignore_index=True)

    # Configure GridOptionsBuilder with JavaScript code
    gb = GridOptionsBuilder.from_dataframe(df_with_totals)

    for column in df.columns[1:]:
        if 'OpenHours' in column:
            gb.configure_column(column, cellStyle=js_code)

    # Allow columns to fill the width and use autoHeight for rows
    gb.configure_grid_options(domLayout= 'normal', autoSizeColumns='allColumns', enableFillHandle=True)

    # Build grid options
    grid_options = gb.build()

    # Set the columnDefs in the grid_options dictionary
    grid_options['columnDefs'] = columnDefs

    # Render the AG-Grid in Streamlit with full width
    try:
        AgGrid(
            df_with_totals,
            gridOptions=grid_options,
            enable_enterprise_modules=True,
            allow_unsafe_jscode=True,  # Allow JavaScript code execution
            fit_columns_on_grid_load=True,  # Automatically fit columns on load
            height=1000,  # Set grid height to 500 pixels
            width='100%',  # Set grid width to 100% of the available space
            theme='streamlit',
            custom_css=custom_css   
        )
    except Exception as ex:
        st.error(f"An error occurred: {ex}")

with tab2:
    if aggregated_data.empty:
        st.warning("No shops found for the selected filter criteria.")
    
    st.markdown(''':green[ **Horas bloqueadas que no se superponen con ninguna cita y tienen turnos configurados*]''')

    # Adjust the pivot table to use BlockedHoursPercentage
    pivot_table_tab2 = aggregated_data.pivot_table(
        index=['Shop[Name]'],
        columns=['date', 'weekday'],
        values='BlockedHoursPercentage',
        aggfunc='mean',
        fill_value=0
    )

    # Flatten the columns for display
    pivot_table_tab2.columns = [f"{col[0]}_{col[1]}" for col in pivot_table_tab2.columns.to_flat_index()]
    pivot_table_tab2_reset = pivot_table_tab2.reset_index()

    # Create the DataFrame for AgGrid
    df_tab2 = pivot_table_tab2_reset

    # Ensure no spaces in field names in df, replacing spaces with underscores or removing them
    df_tab2.columns = [col.replace(' ', '_') for col in df_tab2.columns]

    # JavaScript code for custom cell styling
    js_code = JsCode("""
    function(params) {
        var blockedHoursValue = params.value;

        if (params.data['Shop[Name]'] === 'Total') {
            return {'font-weight': 'bold', 'backgroundColor': '#e0e0e0'};  // Make the total row bold and set a light background color for visibility
        } else if (blockedHoursValue === 0) {
            return {'backgroundColor': '#95cd41', 'color': 'black'};  // Green for BlockedHoursPercentage = 0
        } else if (blockedHoursValue > 0 && blockedHoursValue <= 50) {
            return {'backgroundColor': '#f1b84b', 'color': 'black'};  // Orange for BlockedHoursPercentage between 0 and 50
        } else if (blockedHoursValue > 50) {
            return {'backgroundColor': '#cc0641', 'color': 'white'};  // Red for BlockedHoursPercentage > 50
        } else {
            return null;  // Default style
        }
    }
    """)

    custom_css = {
        ".ag-header-cell": {
            "background-color": "#cc0641 !important",  # Ensure entire cell background changes
            "color": "white !important",
            "font-weight": "bold",
            "padding": "4px"  # Reduce padding to make headers more compact
        },
        ".ag-header-group-cell": {  # Style for merged/group headers
            "background-color": "#cc0641 !important",
            "color": "white !important",
            "font-weight": "bold",
        },
        ".ag-cell": {
            "padding": "2px",  # Reduce padding inside cells to make them more compact
            "font-size": "12px"  # Reduce font size for a more compact look
        },
        ".ag-header": {
            "height": "35px",  # Reduce header height
        },
        ".ag-theme-streamlit .ag-row": {
            "max-height": "30px"  # Adjust max height for rows to be more compact
        },
        ".ag-theme-streamlit .ag-menu-option-text, .ag-theme-streamlit .ag-filter-body-wrapper, .ag-theme-streamlit .ag-input-wrapper, .ag-theme-streamlit .ag-icon": {
            "font-size": "6px !important"  # Reduce font size for filter options and ensure it's applied
        },
        ".ag-theme-streamlit .ag-root-wrapper": {
            "border": "2px solid #cc0641",  # Add outer border with specified color
            "border-radius": "5px"  # Optional: Rounded corners for the outer border
        }
    }

    # Example column definition with flex and resizable properties
    columnDefs_tab2 = [
        {
            "headerName": "Shop Name",
            "field": "Shop[Name]", 
            "resizable": True,
            "flex": 2,  # Adjust flex value to make this column wider
            "minWidth": 150,  # Set a minimum width for columns
            "filter": 'agTextColumnFilter',  # Set filter type to text for shop name
        },
    ]

    # Append dynamic column definitions for BlockedHoursPercentage
    for column in df_tab2.columns[1:]:  # Start from 1 to skip Shop[Name]
        headerName = column.split('_')[0] + ' (' + column.split('_')[1] + ')'
        columnDefs_tab2.append({
            "field": column,
            "headerName": headerName,
            "valueFormatter": "(x > 100 ? 100 : x.toFixed(1)) + ' %'",  # Cap at 100 and add %
            "resizable": True,
            "flex": 1,
            "cellStyle": js_code
        })

    # Calculate totals for BlockedHoursPercentage
    total_row_tab2 = {'Shop[Name]': 'Total'}
    for col in df_tab2.columns[1:]:
        total_row_tab2[col] = df_tab2[col].mean()

    # Convert total_row to DataFrame and append to the original data
    total_df_tab2 = pd.DataFrame(total_row_tab2, index=[0])
    df_with_totals_tab2 = pd.concat([total_df_tab2, df_tab2], ignore_index=True)

    # Configure GridOptionsBuilder with JavaScript code
    gb_tab2 = GridOptionsBuilder.from_dataframe(df_with_totals_tab2)

    for column in df_tab2.columns[1:]:
        gb_tab2.configure_column(column, cellStyle=js_code)

    # Allow columns to fill the width and use autoHeight for rows
    gb_tab2.configure_grid_options(domLayout='normal', autoSizeColumns='allColumns', enableFillHandle=True)

    # Build grid options
    grid_options_tab2 = gb_tab2.build()

    # Set the columnDefs in the grid_options dictionary
    grid_options_tab2['columnDefs'] = columnDefs_tab2

    # Render the AG-Grid in Streamlit
    try:
        AgGrid(
            df_with_totals_tab2,
            gridOptions=grid_options_tab2,
            enable_enterprise_modules=True,
            allow_unsafe_jscode=True,  # Allow JavaScript code execution
            fit_columns_on_grid_load=True,  # Automatically fit columns on load
            height=1000,  # Set grid height to 1000 pixels
            width='100%',  # Set grid width to 100% of the available space
            theme='streamlit',
            custom_css=custom_css   
        )
    except Exception as ex:
        st.error(f"An error occurred: {ex}")
with tab4:
    url1 = "https://amplifongroup.service-now.com/esc?id=esc_dashboard"
    url2 = "https://amplifongroup.service-now.com/sp_amp?id=sc_category_amp&sys_id=9ddfb4d3db96209072ccbb13f3961918"
    if filtered_hcm.empty:
        st.warning("No shops found for the selected filter criteria.")

    st.markdown(f'''
    :green[*Esta vista os puede ayudar para identificar rápidamente dónde cada AP tiene horas asignadas dentro de vuestra área, facilitando el control del negocio en caso de movimientos temporales entre tiendas, por ejemplo. Si vosotros o HRBP necesitais abrir un ticket (HR Ops o Suporte) podeis hacerlo a traves de los siguientes enlances:*]
    - [Abrir Ticket por HR]({url1})
    - [Abrir Ticket por Soporte]({url2})
    ''')

    # Pivot the table for Tab 4
    pivot_table_tab4 = filtered_hcm.pivot_table(
        index=['Resource Name', 'Shop Name'],
        columns='iso_week',
        values=['Duración SF', 'Duración HCM', 'Diferencia de hcm duración'],
        aggfunc='sum',
        fill_value=0
    )
    # Flatten the columns for display
    pivot_table_tab4.columns = [f"Week {col[1]} {col[0]}" for col in pivot_table_tab4.columns.to_flat_index()]
    pivot_table_tab4_reset = pivot_table_tab4.reset_index()
    pivot_table_tab4_reset.columns = [col.replace(' ', '_') for col in pivot_table_tab4_reset.columns]

    # Format all numeric columns to one decimal point
    numeric_columns_in_pivot_tab4 = pivot_table_tab4_reset.select_dtypes(include=['float64', 'int64']).columns
    pivot_table_tab4_reset[numeric_columns_in_pivot_tab4] = pivot_table_tab4_reset[numeric_columns_in_pivot_tab4].round(1)
    # Create the DataFrame for AgGrid
    df_tab4 = pivot_table_tab4_reset
    js_code = JsCode("""
        function(params) {
            // Check if the current row is the totals row by comparing the 'Shop_Name' field
            if (params.data['Resource_Name'] === 'Total') {
                // Apply bold text to the entire totals row and style based on value
                var totalValue = params.value;
                var styles = {'fontWeight': 'bold'};  // Make text bold
                
                if (totalValue === 0) {
                    styles['backgroundColor'] = '#e0e0e0';  // Green for 0
                    styles['color'] = 'white';
                } else if (totalValue > -3 && totalValue < 3) {
                    styles['backgroundColor'] = '#e0e0e0';  // Orange for between -3 and 3
                    styles['color'] = 'black';
                } else {
                    styles['backgroundColor'] = '#e0e0e0';  // Red for outside of -3 and 3
                    styles['color'] = 'black';
                }
                return styles;  // Return styles object
            }

            // Default behavior for non-total rows
            var deltaField = params.colDef.field.replace('Duración_SF', 'Diferencia_de_hcm_duración')
                                                .replace('Duración_HCM', 'Diferencia_de_hcm_duración');
            var deltaValue = params.data[deltaField];
            if (deltaValue === 0) {
                return {'backgroundColor': '#95cd41', 'color': 'black'};  // Green for Delta = 0
            } else if (deltaValue > -3 && deltaValue < 3) {
                return {'backgroundColor': '#f1b84b', 'color': 'black'};  // Orange for Delta between -3 and 3
            } else {
                return {'backgroundColor': '#cc0641', 'color': 'white'};  // Red for Delta < -3 or > 3
            }
        }
    """)

    # Custom CSS for styling the grid, including Week X headers
    custom_css = {
        ".ag-header-cell": {
            "background-color": "#cc0641 !important",  # Set Week X headers to red
            "color": "white !important",
            "font-weight": "bold",
            "padding": "4px"
        },
        ".ag-header-group-cell": {  # Style for merged/group headers
            "background-color": "#cc0641 !important",
            "color": "white !important",
            "font-weight": "bold",
        },
        ".ag-cell": {
            "padding": "2px",
            "font-size": "12px"
        },
        ".ag-theme-streamlit .ag-row": {
            "max-height": "30px"
        },
        ".ag-theme-streamlit .ag-root-wrapper": {
            "border": "2px solid #cc0641",
            "border-radius": "5px"
        }
    }

    # Define the column configuration for "Shop Name"
    columnDefs = [
        {
            "headerName": "Resource Name",
            "field": "Resource_Name",
            "resizable": True,
            "flex": 2,
            "minWidth": 150
        },
        {
            "headerName": "Shop Name",
            "field": "Shop_Name",
            "resizable": True,
            "flex": 2,
            "minWidth": 150
        }

    ]
    # Append dynamic column definitions for each week's SF, HCM, Delta (apply color coding based on Delta value)
    for week in range(start_iso_week, end_iso_week+1):
        columnDefs.append({
            "headerName": f"Week {week}",
            "children": [
                {
                    "field": f"Week_{week}_Duración_SF",
                    "headerName": "SF",
                    "valueFormatter": "x.toFixed(1)",
                    "resizable": True,
                    "flex": 1,
                    "cellStyle": js_code  # Apply the same color coding to totals row
                },
                {
                    "field": f"Week_{week}_Duración_HCM",
                    "headerName": "HCM",
                    "valueFormatter": "x.toFixed(1)",
                    "resizable": True,
                    "flex": 1,
                    "cellStyle": js_code  # Apply the same color coding to totals row
                },
                {
                    "field": f"Week_{week}_Diferencia_de_hcm_duración",
                    "headerName": "Delta",
                    "valueFormatter": "x.toFixed(1)",
                    "resizable": True,
                    "flex": 1,
                    "cellStyle": js_code  # Apply the same color coding to totals row
                }
            ]
        })


    # Calculate totals for numeric columns
    total_row_tab4 = {
        'Resource_Name': 'Total'
    }
    # Only sum columns that actually exist in the DataFrame and apply comma formatting to totals
    for col in numeric_columns_in_pivot_tab4:
        if col in df_tab4.columns:  # Check if column exists before summing
            # Round the totals to 0 decimal places and format with commas
            total_row_tab4[col] = f"{int(df_tab4[col].sum().round(0)):,}"

    # Convert total_row to DataFrame
    total_df_tab4 = pd.DataFrame(total_row_tab4, index=[0])
    df_tab4_with_totals = pd.concat([total_df_tab4, df_tab4], ignore_index=True)  

    # Configure GridOptionsBuilder with JavaScript code
    gb_tab4 = GridOptionsBuilder.from_dataframe(df_tab4_with_totals)
    # Add individual column configurations for conditional formatting
    for column in df_tab4[1:]:
        gb_tab4.configure_column(field=column, cellStyle=js_code)

    # Configure grid options (same as Tab 1)
    gb_tab4.configure_grid_options(domLayout='normal', autoSizeColumns='allColumns', enableFillHandle=True)

    # Build grid options
    grid_options_tab4 = gb_tab4.build()

    # Set the columnDefs in the grid_options dictionary
    grid_options_tab4['columnDefs'] = columnDefs

    # Render the AG-Grid in Streamlit with full width and custom styling
    AgGrid(
        df_tab4_with_totals,
        gridOptions=grid_options_tab4,
        enable_enterprise_modules=True,
        allow_unsafe_jscode=True,
        fit_columns_on_grid_load=True,
        height=1000,
        width='100%',
        theme='streamlit',
        custom_css=custom_css  # Apply custom CSS
    )
with tab3:
    st.markdown(''':green[ **Las horas reales trabajadas se suman a todas las tiendas por audiólogo. La tienda que se muestra es solo una de las tiendas en las que está registrado el audiólogo.*]''')
    st.markdown(''':green[ **Los datos de entrada y salida están disponibles a partir del 3 de septiembre. Los turnos de SF se ajustan por bloqueos/ausencias.*]''')
    st.markdown(''':green[ **Todos los audiólogos que no hayan registrado su salida estarán marcados como NC (No completo) para el día.*]''')

    if filtered_clock.empty:
        st.warning("No filtre por tienda. Las horas reales trabajadas se suman a todas las tiendas por empleado.  ")

    # Pivot the table for Tab 3
    def custom_agg_hours(series):
        # If all values are 'NC', return 'NC'
        if (series == 'NC').all():
            return 'NC'
        # Convert 'NC' to 0 for summation purposes
        numeric_series = pd.to_numeric(series.replace('NC', 0), errors='coerce').fillna(0)
        return numeric_series.sum()

    # Pivot the table with custom aggregation function for 'hours_worked'
    pivot_table_tab3 = filtered_clock.pivot_table(
        index=['Resource Name', 'Shop[Name]'],
        columns=['Date', 'weekday'],
        values=['hours_worked', 'ShiftDurationHoursAdjusted', 'Diferencia de act duración'],
        aggfunc={
            'hours_worked': custom_agg_hours,
            'ShiftDurationHoursAdjusted': 'sum',
            'Diferencia de act duración': 'sum'
        },
        fill_value=0
    )
    # Flatten the columns (removing 'Day' and aligning with the other tab)
    pivot_table_tab3.columns = [
        f"{col[0]}_{col[1].strftime('%Y-%m-%d')} {col[2]}" for col in pivot_table_tab3.columns.to_flat_index()
    ] 
    pivot_table_tab3_reset = pivot_table_tab3.reset_index()
    # Format all numeric columns to one decimal point
    numeric_columns_in_pivot_tab3 = [col for col in pivot_table_tab3_reset.columns if 'hours_worked' in col or 'ShiftDurationHoursAdjusted' in col or 'Diferencia de act duración' in col]
    pivot_table_tab3_reset[numeric_columns_in_pivot_tab3] = pivot_table_tab3_reset[numeric_columns_in_pivot_tab3].round(1)
    # Create the DataFrame for AgGrid
    df_tab3 = pivot_table_tab3_reset
    # Ensure no spaces in field names, replace spaces with underscores
    df_tab3.columns = [col.replace(' ', '_') for col in df_tab3.columns]

    js_code = JsCode("""
        function(params) {
            // Check if the current row is the totals row by comparing the 'Shop_Name' field
            if (params.data['Resource_Name'] === 'Total') {
                var totalValue = params.value;
                var styles = {'fontWeight': 'bold'};  // Make text bold
                
                if (totalValue === 0) {
                    styles['backgroundColor'] = '#e0e0e0';
                    styles['color'] = 'white';
                } else if (totalValue > -3 && totalValue < 3) {
                    styles['backgroundColor'] = '#e0e0e0';
                    styles['color'] = 'black';
                } else {
                    styles['backgroundColor'] = '#e0e0e0';
                    styles['color'] = 'black';
                }
                return styles;
            }

            var deltaField = params.colDef.field.replace('ShiftDurationHoursAdjusted', 'Diferencia_de_act_duración')
                                                .replace('hours_worked', 'Diferencia_de_act_duración');
            var deltaValue = params.data[deltaField];
            if (deltaValue === 0) {
                return {'backgroundColor': '#95cd41', 'color': 'black'};
            } else if (deltaValue > -3 && deltaValue < 3) {
                return {'backgroundColor': '#f1b84b', 'color': 'black'};
            } else {
                return {'backgroundColor': '#cc0641', 'color': 'white'};
            }
        }
    """)

    # Custom CSS for styling the grid, including day headers
    custom_css = {
        ".ag-header-cell": {
            "background-color": "#cc0641 !important",
            "color": "white !important",
            "font-weight": "bold",
            "padding": "4px"
        },
        ".ag-header-group-cell": {
            "background-color": "#cc0641 !important",
            "color": "white !important",
            "font-weight": "bold",
        },
        ".ag-cell": {
            "padding": "2px",
            "font-size": "12px"
        },
        ".ag-theme-streamlit .ag-row": {
            "max-height": "30px"
        },
        ".ag-theme-streamlit .ag-root-wrapper": {
            "border": "2px solid #cc0641",
            "border-radius": "5px"
        }
    }

    columnDefs = [
        {
            "headerName": "Resource Name",
            "field": "Resource_Name",
            "resizable": True,
            "flex": 2,
            "minWidth": 150
        },
        {
            "headerName": "Shop Name",
            "field": "Shop[Name]",
            "resizable": True,
            "flex": 2,
            "minWidth": 150
        }
    ]


    # Append dynamic column definitions for SF (ShiftDurationHoursAdjusted), ACT (hours_worked), and Delta (Diferencia de act duración)
    for column in df_tab3.columns[2:]:  # Start from 2 to skip Resource_Name and Shop[Name]
        if 'ShiftDurationHoursAdjusted' in column:
            headerName = column.split('_')[1] + ' (' + column.split('_')[2] + ')'
            columnDefs.append({
                "headerName": headerName,
                "children": [
                    {
                        "field": column,
                        "headerName": "SF",  # ShiftDurationHoursAdjusted
                        "valueFormatter": "x.toFixed(1)",
                        "resizable": True,
                        "flex": 1,
                        "cellStyle": js_code
                    },
                    {
                        "field": column.replace('ShiftDurationHoursAdjusted', 'hours_worked'),
                        "headerName": "ACT",  # hours_worked
                        "valueFormatter": "x.toFixed(1)",
                        "resizable": True,
                        "flex": 1,
                        "cellStyle": js_code
                    },
                    {
                        "field": column.replace('ShiftDurationHoursAdjusted', 'Diferencia_de_act_duración'),
                        "headerName": "Delta",  # Diferencia de act duración
                        "valueFormatter": "x.toFixed(1)",
                        "resizable": True,
                        "flex": 1,
                        "cellStyle": js_code
                    }
                ]
            })

    # Calculate totals for numeric columns
    total_row_tab3 = {
        'Resource_Name': 'Total'
    }
    for col in numeric_columns_in_pivot_tab3:
        col = col.replace(' ', '_') 
        if col in df_tab3.columns:  
            # Convert column to numeric, ignoring errors to handle 'NC'
            numeric_values = pd.to_numeric(df_tab3[col], errors='coerce')
            total_value = numeric_values.sum(min_count=1)  # Sum ignoring non-numeric, use `min_count=1` to handle empty
            if pd.isna(total_value):
                total_value = 'NC' 
            else:
                total_value = f"{int(total_value.round(0)):,}"  
            total_row_tab3[col] = total_value
    # Convert total_row to DataFrame
    total_df_tab3 = pd.DataFrame(total_row_tab3, index=[0])
    df_tab3_with_totals = pd.concat([total_df_tab3, df_tab3], ignore_index=True)  

    # Configure GridOptionsBuilder with JavaScript code
    gb_tab3 = GridOptionsBuilder.from_dataframe(df_tab3_with_totals)
    # Add individual column configurations for conditional formatting
    for column in df_tab3[1:]:
        gb_tab3.configure_column(field=column, cellStyle=js_code)

    gb_tab3.configure_grid_options(domLayout='normal', autoSizeColumns='allColumns', enableFillHandle=True)

    # Build grid options
    grid_options_tab3 = gb_tab3.build()

    # Set the columnDefs in the grid_options dictionary
    grid_options_tab3['columnDefs'] = columnDefs

    # Render the AG-Grid in Streamlit with full width and custom styling
    AgGrid(
        df_tab3_with_totals,
        gridOptions=grid_options_tab3,
        enable_enterprise_modules=True,
        allow_unsafe_jscode=True,
        fit_columns_on_grid_load=True,
        height=1000,
        width='100%',
        theme='streamlit',
        custom_css=custom_css
    )



with tab5:

    # Warning messages for empty data
    if weekly_shift_slots.empty:
        st.warning("No shops found for the selected filter criteria.")
    if weekly_shift_slots_yesterday.empty:
        st.warning("No shops found for the selected filter criteria.")

    # Define the date range for filtered_clock_noiso
    start_date_act = pd.Timestamp('2024-10-03')
    end_date_act = pd.Timestamp(datetime.now().date() - timedelta(days=1))

    # Column layout
    st.markdown("### Overview")
    cols = st.columns(2)

    # HCM vs SF Chart and Top Shops
    with cols[0]:
        # Step 1: Group filtered_hcm by 'iso_week' and sum 'Diferencia de hcm duración'
        hcm_weekly_diff = filtered_hcm.groupby('iso_week').agg(
            total_diff=('Diferencia de hcm duración', 'sum')
        ).reset_index()

        # Create HCM vs SF line chart
        fig_diff = px.line(
            hcm_weekly_diff,
            x='iso_week', 
            y='total_diff',
            labels={'iso_week': 'ISO Week', 'total_diff': 'Total Difference (HCM vs SF)'},
            title="HCM vs SF Duration Difference Over Weeks",
            markers=True
        )
        fig_diff.update_layout(
            xaxis_title="Week Number",
            yaxis_title="Total Difference",
            hovermode="x unified"
        )
        fig_diff.update_traces(
            hovertemplate='ISO Week: %{x}<br>Total Difference: %{y:.2f}'
        )
        cols[0].plotly_chart(fig_diff)

        # Get top 3 shops with the largest 'Diferencia de hcm duración'
        top_shops_hcm = filtered_hcm.groupby('Shop Name').agg(
            total_diff=('Diferencia de hcm duración', 'sum')
        ).nlargest(3, 'total_diff').reset_index()

        # Display top 3 shops side-by-side under the HCM vs SF chart
        hcm_cols = st.columns(3)
        for i, row in enumerate(top_shops_hcm.iterrows()):
            with hcm_cols[i]:
                st.markdown(f"""
                <div style="border: 2px solid #1f77b4; border-radius: 10px; padding: 10px; margin: 5px 0; background-color: #f9f9f9; text-align: center;">
                    <strong>Shop {i + 1}: {row[1]['Shop Name']}</strong><br>
                    Total Difference: <span style="color: #1f77b4; font-weight: bold;">{row[1]['total_diff']:.2f}</span>
                </div>
                """, unsafe_allow_html=True)


    # ACT vs SF Chart and Top Shops
    with cols[1]:
        # Filter date range for ACT vs SF data
        filtered_clock_date_range = filtered_clock_noiso[
            (filtered_clock_noiso['Date'] >= start_date_act) & (filtered_clock_noiso['Date'] <= end_date_act)
        ]

        # Group by Date and sum 'Diferencia de act duración'
        acd_weekly_diff = filtered_clock_date_range.groupby('Date').agg(
            total_diff=('Diferencia de act duración', 'sum')
        ).reset_index()

        # Create ACT vs SF line chart
        fig_acd_diff = px.line(
            acd_weekly_diff,
            x='Date',
            y='total_diff',
            labels={'Date': 'Day', 'total_diff': 'Total Difference (ACT vs SF)'},
            title="ACT vs SF Duration Difference Over Days",
            markers=True
        )
        fig_acd_diff.update_layout(
            xaxis_title="Day",
            yaxis_title="Total Difference",
            hovermode="x unified"
        )
        fig_acd_diff.update_traces(
            hovertemplate='Day: %{x}<br>Total Difference: %{y:.2f}',
            line_color='red'
        )
        cols[1].plotly_chart(fig_acd_diff)

        # Get top 3 shops with the largest 'Diferencia de act duración'
        top_shops_acd = filtered_clock_date_range.groupby('Shop[Name]').agg(
            total_diff=('Diferencia de act duración', 'sum')
        ).nlargest(3, 'total_diff').reset_index()

        # Display top 3 shops side-by-side under the ACT vs SF chart
        acd_cols = st.columns(3)
        for i, row in enumerate(top_shops_acd.iterrows()):
            with acd_cols[i]:
                st.markdown(f"""
                <div style="border: 2px solid #d62728; border-radius: 10px; padding: 10px; margin: 5px 0; background-color: #f9f9f9; text-align: center;">
                    <strong>Shop {i + 1}: {row[1]['Shop[Name]']}</strong><br>
                    Total Difference: <span style="color: #d62728; font-weight: bold;">{row[1]['total_diff']:.2f}</span>
                </div>
                """, unsafe_allow_html=True)


    # Step 1: Aggregating summary_tab_data by iso_week to get total hours per week
    weekly_aggregated = weekly_shift_slots.groupby('iso_week').agg(
        TotalHours=('TotalHours', 'sum'),
        BlockedHours=('BlockedHours', 'sum'),
        AvailableHours=('AvailableHours', 'sum'),
        BookedHours=('BookedHours', 'sum'),
        OpenHours=('OpenHours', 'sum')
    ).reset_index()

    weekly_aggregated = weekly_aggregated.fillna(0)

    # Create an interactive time series graph with Plotly
    fig = px.line(
        weekly_aggregated, 
        x='iso_week',  
        y=['TotalHours', 'BlockedHours', 'AvailableHours', 'BookedHours', 'OpenHours'],
        labels={'iso_week': 'ISO Week', 'value': 'Hours'},  
        title="Weekly Hours Overview",
        markers=True  
    )

    # Customize the layout for better readability
    fig.update_layout(
        xaxis_title="Week Number",  
        yaxis_title="Hours", 
        hovermode="x unified"  
    )

    fig.update_traces(
        hovertemplate='%{y:,.0f} Hours<br>ISO Week: %{x}'
    )

    # Display the plotly graph in Streamlit
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("### Tiendas que necesitan atención urgente (Para Hoy)")

    # Get today's date
    today = datetime.now().date()

    # Convert the 'date' column to datetime
    weekly_shift_slots['date'] = pd.to_datetime(weekly_shift_slots['date'], format='%Y-%m-%d')

    # Step 1: Filter the dataset to include necessary columns for today
    filtered_data = weekly_shift_slots[weekly_shift_slots['date'].dt.date == today][['Region', 'Shop[Name]', 'BlockedHoursPercentage']].copy()

    # Sort by region and by highest BlockedHoursPercentage
    top_shops_by_region = (
        filtered_data.sort_values(by='BlockedHoursPercentage', ascending=False)
        .groupby('Region')
        .head(4)  # Get top 4 shops per region
    )

    # Get unique regions
    regions = top_shops_by_region['Region'].unique()

    # Custom CSS for styling the boxes and the shop list
    st.markdown("""
        <style>
        .custom-box {
            border: 2px solid #cc0641;
            border-radius: 10px;
            padding: 15px;
            margin-bottom: 10px;
            background-color: #f9f9f9;
            text-align: center;
            width: 100%;
            box-sizing: border-box;
        }
        .custom-box h5 {
            color: #cc0641;
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 10px;
            text-transform: uppercase;
        }
        .shop-list {
            margin-top: 10px;
            text-align: left;
            line-height: 1.6;  /* Increased line height for readability */
        }
        .shop-list p {
            font-size: 14px;
            margin: 0;
        }
        .shop-list strong {
            color: #333;
        }
        </style>
    """, unsafe_allow_html=True)

    # Create 4 columns to display 4 regions in one row
    cols = st.columns(4)

    # Iterate over regions and display them in the respective column
    for i, region in enumerate(regions):
        with cols[i % 4]:  # Ensure 4 regions per row
            # Start the custom box div and shop list inside it
            top_shops = top_shops_by_region[top_shops_by_region['Region'] == region]

            shop_list = ""
            for index, row in top_shops.iterrows():
                shop_name = row['Shop[Name]']
                blocked_pct = row['BlockedHoursPercentage']
                
                # Append shop details to the list
                shop_list += f"<p>- <strong>{shop_name}</strong>: {blocked_pct:,.0f}% Blocked Hours</p>"

            # Render both the header and the shop list inside the box
            st.markdown(f"""
            <div class="custom-box">
                <h5>{region}</h5>
                <div class="shop-list">
                    {shop_list}
                </div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("### Tiendas Cerradas (Para Hoy)")

    # Step 1: Filter the dataset for closed shops (TotalHours = 0 or BlockedHoursPercentage = 100%)
    closed_shops_data = weekly_shift_slots[
        (weekly_shift_slots['date'].dt.date == today) &
        ((weekly_shift_slots['TotalHours'] == 0) | (weekly_shift_slots['BlockedHoursPercentage'] == 100))
    ][['Region', 'Shop[Name]', 'TotalHours', 'BlockedHoursPercentage']].copy()

    # Sort by region
    closed_shops_data = closed_shops_data.sort_values(by='Region')

    # Get unique regions with closed shops
    regions_with_closed_shops = closed_shops_data['Region'].unique()

    # Custom CSS for styling the boxes and the shop list (same style as before)
    st.markdown("""
        <style>
        .custom-box {
            border: 2px solid #cc0641;
            border-radius: 10px;
            padding: 15px;
            margin-bottom: 10px;
            background-color: #f9f9f9;
            text-align: center;
            width: 100%;
            box-sizing: border-box;
        }
        .custom-box h5 {
            color: #cc0641;
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 10px;
            text-transform: uppercase;
        }
        .shop-list {
            margin-top: 10px;
            text-align: left;
            line-height: 1.6;  /* Increased line height for readability */
        }
        .shop-list p {
            font-size: 14px;
            margin: 0;
        }
        .shop-list strong {
            color: #333;
        }
        </style>
    """, unsafe_allow_html=True)

    # Split regions into groups of 4 (for displaying in 4 columns per row)
    for start_idx in range(0, len(regions_with_closed_shops), 4):
        cols_closed = st.columns(4)  # Create 4 columns for this row
        
        # Display up to 4 regions in the current row
        for i, region in enumerate(regions_with_closed_shops[start_idx:start_idx + 4]):
            with cols_closed[i]:  # Ensure 4 regions per row
                # Filter the data for the current region
                closed_shops = closed_shops_data[closed_shops_data['Region'] == region]

                shop_list = ""
                for index, row in closed_shops.iterrows():
                    shop_name = row['Shop[Name]']
                    total_hours = row['TotalHours']
                    blocked_pct = row['BlockedHoursPercentage']

                    # Append shop details to the list
                    shop_list += f"<p>- <strong>{shop_name}</strong>: {total_hours} Hours, {blocked_pct:,.0f}% Blocked</p>"

                # Render both the header and the shop list inside the box
                st.markdown(f"""
                <div class="custom-box">
                    <h5>{region}</h5>
                    <div class="shop-list">
                        {shop_list}
                    </div>
                </div>
                """, unsafe_allow_html=True)
                    
# Format the month_start_date as a string like "October 1"

with tab6:
    if weekly_shift_slots.empty:
        st.warning("No shops found for the selected filter criteria.")
    if weekly_shift_slots_yesterday.empty:
        st.warning("No shops found for the selected filter criteria.")
    
    formatted_month_start_date = month_start_date.strftime("%B %#d").lstrip('0')

    st.markdown("### Weekly Overview")
    # Add a new selectbox for comparison options
    comparison_options = [f'Comparison with start of the month ({formatted_month_start_date})', 'Comparison with yesterday']
    selected_comparison = st.selectbox('Select Comparison:', comparison_options)
    
    metric_options = ['Shift Hours % change', 'Blocked Hours % change', 'Booked Hours % change', 'Open Hours % change']
    selected_metric = st.selectbox('Select Metric:', metric_options)
    metric_map = {
        'Shift Hours % change': 'TotalHours',
        'Blocked Hours % change': 'BlockedHours',
        'Booked Hours % change': 'BookedHours',
        'Open Hours % change': 'OpenHours'
    }

    # Get the column associated with the selected metric
    metric_column = metric_map[selected_metric]

    weekly_shift_slots = weekly_shift_slots[(weekly_shift_slots['date'] >= month_start_date) & (weekly_shift_slots['date'] <= month_end_date)].copy()
    weekly_shift_slots_yesterday = weekly_shift_slots_yesterday[(weekly_shift_slots_yesterday['date'] >= month_start_date) & (weekly_shift_slots_yesterday['date'] <= month_end_date)].copy()
    weekly_shift_sep6 = weekly_shift_sep6[(weekly_shift_sep6['date'] >= month_start_date) & (weekly_shift_sep6['date'] <= month_end_date)].copy()        

    # Pivot the table for Tab 4
    weekly_aggregated = weekly_shift_slots.pivot_table(
        index='Region',
        columns='iso_week',
        values=f'{metric_column}',
        aggfunc='sum',
        fill_value=0
    )
    weekly_aggregated = weekly_aggregated.fillna(0)
    weekly_aggregated = weekly_aggregated.reset_index()

    # Step 1: Aggregating summary_tab_data by iso_week to get total hours per week
    weekly_aggregated_yesterday = weekly_shift_slots_yesterday.pivot_table(
        index='Region',
        columns='iso_week',
        values=f'{metric_column}',
        aggfunc='sum',
        fill_value=0
    )
    weekly_aggregated_yesterday = weekly_aggregated_yesterday.fillna(0)
    weekly_aggregated_yesterday = weekly_aggregated_yesterday.reset_index()

    # Step 1: Aggregating summary_tab_data by iso_week to get total hours per week
    weekly_aggregated_sep6 = weekly_shift_sep6.pivot_table(
        index='Region',
        columns='iso_week',
        values=f'{metric_column}',
        aggfunc='sum',
        fill_value=0
    )
    weekly_aggregated_sep6 = weekly_aggregated_sep6.fillna(0)
    weekly_aggregated_sep6 = weekly_aggregated_sep6.reset_index()
    
    # Adjust the logic to use different dataframes based on the comparison selection
    if selected_comparison == f'Comparison with start of the month ({formatted_month_start_date})':
        comparison_df = weekly_aggregated_sep6  # Assuming Sep 6 data is stored in weekly_aggregated_yesterday
        comparison_label = formatted_month_start_date
    else:
        comparison_df = weekly_aggregated_yesterday  # Assuming yesterday's data is in weekly_aggregated_yesterday2
        comparison_label = 'Yesterday'

    # Step 1: Melt the data to convert wide to long format
    weekly_aggregated_melted = pd.melt(
        weekly_aggregated, 
        id_vars=['Region'], 
        var_name='iso_week', 
        value_name=f'{metric_column}_today'
    )
    comparison_melted = pd.melt(
        comparison_df, 
        id_vars=['Region'], 
        var_name='iso_week', 
        value_name=f'{metric_column}_comparison'
    )

    # Step 2: Merge the two dataframes on Region and iso_week
    merged_data = pd.merge(
        weekly_aggregated_melted, 
        comparison_melted, 
        on=['Region', 'iso_week'], 
        how='inner'
    )

    merged_grouped = merged_data.groupby(['Region', 'iso_week']).agg(
        {
            f'{metric_column}_today': 'sum',  # Sum today's metric values
            f'{metric_column}_comparison': 'sum'  # Sum yesterday's metric values
        }
    ).reset_index()
    weekly_totals = merged_grouped.groupby('iso_week').agg(
        {f'{metric_column}_today': 'sum', f'{metric_column}_comparison': 'sum'}
    ).reset_index()

    # Step 3: Calculate monthly totals (sum over all weeks)
    monthly_total_today = weekly_totals[f'{metric_column}_today'].sum()
    monthly_total_comparison = weekly_totals[f'{metric_column}_comparison'].sum()

    # Step 7: Add the month column for each region based on the aggregated data
    region_month_totals = merged_grouped.groupby('Region').agg(
        {f'{metric_column}_today': 'sum', f'{metric_column}_comparison': 'sum'}
    ).reset_index()

    def calculate_percentage_change(today_value, yesterday_value):
        if yesterday_value == 0:
            return 0
        return ((today_value - yesterday_value) / abs(yesterday_value)) * 100
    
    # Step 2: Calculate percentage change for each week based on the total values
    weekly_totals['total_percentage_change'] = weekly_totals.apply(
        lambda row: calculate_percentage_change(row[f'{metric_column}_today'], row[f'{metric_column}_comparison']), axis=1
    )
    # Step 3: Calculate percentage change
    merged_data['percentage_change'] = merged_data.apply(
        lambda row: calculate_percentage_change(row[f'{metric_column}_today'], row[f'{metric_column}_comparison']), axis=1
    )
    monthly_percentage_change = calculate_percentage_change(monthly_total_today, monthly_total_comparison)
    # Calculate percentage change for the month for each region
    region_month_totals['month_percentage_change'] = region_month_totals.apply(
        lambda row: calculate_percentage_change(row[f'{metric_column}_today'], row[f'{metric_column}_comparison']), axis=1
    )
    # Step 4: Optional sorting or filtering
    merged_data_sorted = merged_data.sort_values(by=['Region', 'iso_week'])
    merged_pivot = merged_data_sorted.pivot_table(
        index='Region',
        columns='iso_week',
        values='percentage_change',
    ).reset_index()
    # Append the 'Month Total' column to the pivot table for each region
    merged_pivot['Month Total'] = merged_pivot['Region'].map(
        dict(zip(region_month_totals['Region'], region_month_totals['month_percentage_change']))
    )

    # Fill NaN values for the 'Total' row in the 'Month Total' column
    merged_pivot['Month Total'].fillna(monthly_percentage_change, inplace=True)
    # Step 5: Append the total row to the pivot table
    total_row = pd.DataFrame(
        [['Total'] + weekly_totals['total_percentage_change'].tolist() + [monthly_percentage_change]], 
        columns=merged_pivot.columns.tolist()  
    )
    merged_pivot = pd.concat([merged_pivot, total_row], ignore_index=True)

    # Optional: Rename the columns for better display (e.g., 'Week 36', 'Week 37', etc.)
    merged_pivot.columns = ['Region'] + [f"Week {int(col)}" if col != 'Month Total' else col for col in merged_pivot.columns[1:]]
    
    # Custom CSS for the table (same as Tab 2, adjusted if needed)
    custom_css_tab6 = {
        ".ag-header-cell": {
            "background-color": "#cc0641 !important",  
            "color": "white !important",
            "font-weight": "bold",
            "padding": "4px"
        },
        ".ag-cell": {
            "padding": "2px",
            "font-size": "12px"
        },
        ".ag-header": {
            "height": "35px"
        },
        ".ag-theme-streamlit .ag-row": {
            "max-height": "30px"
        },
        ".ag-theme-streamlit .ag-root-wrapper": {
            "border": "2px solid #cc0641",
            "border-radius": "5px"
        }
    }
    if selected_metric == 'Blocked Hours % change':
        color_coding_js = JsCode("""
        function(params) {
            var value = params.value;

            if (params.data['Region'] === 'Total') {
                return {'font-weight': 'bold', 'backgroundColor': '#e0e0e0'};  // Grey background for total row
            } else {
                // Color coding for Blocked Hours % Change
                if (value <= 0) {
                    return {'backgroundColor': '#95cd41', 'color': 'black'};  // Green for negative or zero BlockedHours % Change
                } else if (value > 0) {
                    return {'backgroundColor': '#cc0641', 'color': 'white'};  // Red for positive BlockedHours % Change
                }
                return null;  // Default styling for other values
            }
        }
        """)
    elif selected_metric == 'Shift Hours % change':
        color_coding_js = JsCode("""
        function(params) {
            var value = params.value;

            if (params.data['Region'] === 'Total') {
                return {'font-weight': 'bold', 'backgroundColor': '#e0e0e0'};  // Grey background for total row
            } else {
                // Color coding for Total Hours % Change (reverse logic)
                if (value < 0) {
                    return {'backgroundColor': '#cc0641', 'color': 'white'};  // Red for positive TotalHours % Change
                } else if (value >= 0) {
                    return {'backgroundColor': '#95cd41', 'color': 'black'};  // Green for negative TotalHours % Change
                }
                return null;  // Default styling for other values
            }
        }
        """)
    elif selected_metric == 'Booked Hours % change':
        color_coding_js = JsCode("""
        function(params) {
            var value = params.value;

            if (params.data['Region'] === 'Total') {
                return {'font-weight': 'bold', 'backgroundColor': '#e0e0e0'};  // Grey background for total row
            } else {
                // Color coding for Total Hours % Change
                if (value < 0) {
                    return {'backgroundColor': '#cc0641', 'color': 'white'};  // Red for positive TotalHours % Change
                } else if (value >= 0) {
                    return {'backgroundColor': '#95cd41', 'color': 'black'};  // Green for negative TotalHours % Change
                }
                return null;  // Default styling for other values
            }
        }
        """)
    elif selected_metric == 'Open Hours % change':
        color_coding_js = JsCode("""
        function(params) {
            var value = params.value;

            if (params.data['Region'] === 'Total') {
                return {'font-weight': 'bold', 'backgroundColor': '#e0e0e0'};  // Grey background for total row
            } else {
                // Color coding for Open Hours % Change
                if (value <= 0) {
                    return {'backgroundColor': '#95cd41', 'color': 'black'};  // Green for negative or zero OpenHours % Change
                } else if (value > 0) {
                    return {'backgroundColor': '#cc0641', 'color': 'white'};  // Red for positive OpenHours % Change
                }
                return null;  // Default styling for other values
            }
        }
        """)


    # Create column definitions for percentage change table
    columnDefs_tab6_pct_change = [{"field": 'Region', "headerName": "Region", "resizable": True, "flex": 1}]
    for column in merged_pivot.columns[1:]:
        columnDefs_tab6_pct_change.append({
            "field": column,
            "headerName": column,  # Use 'Week {iso_week}' as the header
            "valueFormatter": "(x !== null && x !== undefined ? x.toFixed(1) + ' %' : '0 %')", 
            "resizable": True,
            "flex": 1,
            "cellStyle": color_coding_js  # Apply the specific function for the percentage table
        })

    # GridOptionsBuilder for percentage change table
    gb_tab6_pct_change = GridOptionsBuilder.from_dataframe(merged_pivot)
    for col in merged_pivot.columns:  # For the percentage change table
        gb_tab6_pct_change.configure_column(col, cellStyle=color_coding_js)

    # Grid options for auto-sizing and responsive layout
    gb_tab6_pct_change.configure_grid_options(domLayout='normal', autoSizeColumns='allColumns', enableFillHandle=True)

    # Build grid options
    grid_options_tab6_pct_change = gb_tab6_pct_change.build()

    # Add custom column definitions to grid options
    grid_options_tab6_pct_change['columnDefs'] = columnDefs_tab6_pct_change

    # Render pivot_pct_change table using AgGrid
    st.markdown(f"### {selected_metric}: Today vs {comparison_label}")

    AgGrid(
        merged_pivot,
        gridOptions=grid_options_tab6_pct_change,
        enable_enterprise_modules=True,
        allow_unsafe_jscode=True,  # Allow JavaScript code execution
        fit_columns_on_grid_load=True,
        height=187,  # Set grid height for percentage change table
        width='100%',
        theme='streamlit',
        custom_css=custom_css_tab6
      )

    # Aggregate totals for each iso_week across all regions
    merged_grouped_total = merged_grouped.groupby('iso_week').agg(
        {f'{metric_column}_today': 'sum', f'{metric_column}_comparison': 'sum'}
    ).reset_index()

    # Create a new row for the monthly total without formatting for calculation purposes
    monthly_totals_row = pd.DataFrame({
        'iso_week': ['Month Total'],  # Label for the new row
        f'{metric_column}_today': [monthly_total_today],  # Monthly total for today
        f'{metric_column}_comparison': [monthly_total_comparison]  # Monthly total for yesterday
    })

    # Append the new row to the merged_grouped_total DataFrame
    merged_grouped_total = pd.concat([merged_grouped_total, monthly_totals_row], ignore_index=True)

    # Format numbers with commas and round them to integers
    merged_grouped_total[f'{metric_column}_today'] = merged_grouped_total[f'{metric_column}_today'].round(0).apply(lambda x: f"{int(x):,}")
    merged_grouped_total[f'{metric_column}_comparison'] = merged_grouped_total[f'{metric_column}_comparison'].round(0).apply(lambda x: f"{int(x):,}")
    # Create an interactive bar chart using Plotly
    fig = go.Figure()

    # Adjust the comparison label dynamically
    comparison_label = formatted_month_start_date if selected_comparison == f'Comparison with start of the month ({formatted_month_start_date})' else 'Yesterday'

    # Add bars for 'today' values
    fig.add_trace(go.Bar(
        x=merged_grouped_total['iso_week'],
        y=merged_grouped_total[f'{metric_column}_today'].apply(lambda x: int(x.replace(',', ''))),  # Plot the numeric values
        name='Today',
        marker_color='#cc0641',  # Use the custom color for 'Today'
        text=merged_grouped_total[f'{metric_column}_today'],  # Show the formatted values
        textposition='auto'
    ))

    # Add bars for 'comparison' values with a dynamic label (either 'Yesterday' or month_start_date)
    fig.add_trace(go.Bar(
        x=merged_grouped_total['iso_week'],
        y=merged_grouped_total[f'{metric_column}_comparison'].apply(lambda x: int(x.replace(',', ''))),  # Plot the numeric values
        name=comparison_label,  # Dynamic label for the comparison column
        marker_color='#f1b84b',  # Lighter shade for the comparison
        text=merged_grouped_total[f'{metric_column}_comparison'],  # Show the formatted values
        textposition='auto'
    ))

    # Customize layout
    fig.update_layout(
        title=f'Comparison of {metric_column} for Today vs {comparison_label} (Aggregated Across Regions)',  # Adjust the title dynamically
        xaxis=dict(
            title='ISO Week / Month', 
            type='category',
            categoryorder='array',  # Order the x-axis categories manually
            categoryarray=list(merged_grouped_total['iso_week'])  # Set the correct order: weeks followed by "Month Total"
        ),
        yaxis=dict(title=f'{metric_column}'),
        barmode='group',  # Group the bars for today and the comparison side-by-side
        bargap=0.2,  # Set gap between bars
        bargroupgap=0.1,  # Set gap between groups
        legend_title="Metric",
        font=dict(size=12),  # Adjust font size for better readability
    )

    # Render the plot in Streamlit
    st.plotly_chart(fig)
