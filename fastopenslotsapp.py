import sys
import streamlit as st
import pandas as pd
import pytz
from datetime import datetime, timedelta
from st_aggrid import GridOptionsBuilder, AgGrid, JsCode
from st_aggrid.shared import GridUpdateMode, DataReturnMode, ColumnsAutoSizeMode, AgGridTheme, ExcelExportMode
from st_aggrid.AgGridReturn import AgGridReturn
import json
import os

@st.cache_data
def load_excel(file_path, usecols=None, file_mod_time=None, **kwargs):
    # Load specific columns if usecols is provided to reduce memory usage
    return pd.read_excel(file_path, usecols=usecols, **kwargs)

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

file_mod_time = os.path.getmtime('shiftslots.xlsx')

shift_slots = load_excel('shiftslots.xlsx', file_mod_time=file_mod_time)

file_mod_time1 = os.path.getmtime('hcpshiftslots.xlsx')

hcp_shift_slots = load_excel('hcpshiftslots.xlsx', file_mod_time=file_mod_time1)

file_mod_time2 = os.path.getmtime('hcm_sf_merged.xlsx')

hcm = load_excel('hcm_sf_merged.xlsx', file_mod_time=file_mod_time2)

# Sidebar filters (all converted to single-selection using selectbox)
iso_week_filter = st.sidebar.selectbox('Select ISO Week', sorted(shift_slots['iso_week'].unique()))

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

# Apply the filters for the other datasets
filtered_data = filter_data(shift_slots, iso_week_filter, selected_region, selected_area, selected_shop, 'iso_week')
filtered_hcp_shift_slots = filter_data(hcp_shift_slots, iso_week_filter, selected_region, selected_area, selected_shop, 'iso_week')

# Apply the filters to HCM data (without iso_week filter)
filtered_hcm = filter_hcm_data(hcm, selected_region, selected_area, selected_shop)

# Check if filtered data is empty after applying the filters
if filtered_data.empty:
    st.warning("No shops found for the selected filter criteria.")
# Check if filtered data is empty after applying the filters
if filtered_hcp_shift_slots.empty:
    st.warning("No shops found for the selected filter criteria.")
if filtered_hcm.empty:
    st.warning("No shops found for the selected filter criteria.")

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
st.sidebar.markdown(f"<div class='sidebar-stats-box'>Open hours for the selected week: {open_hours_this_week:.2f}</div>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div class='sidebar-stats-box'>Change from last week: {change_from_last_week:.2f}%</div>", unsafe_allow_html=True)
st.sidebar.markdown(f"<div class='sidebar-stats-box'>Open hours for month to go: {open_hours_month_to_go:.2f}</div>", unsafe_allow_html=True)
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

tab1, tab2, tab3, tab4 = st.tabs(["Open Hours / Total Hours", "Blocked Hours Percentage", "MTD/MTG", "HCM vs SF"])

with tab1:        
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
        total_row[col] = df[col].sum()

    # Convert total_row to DataFrame
    total_df = pd.DataFrame(total_row, index=[0])

    df_with_totals = pd.concat([df, total_df], ignore_index=True)

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
    df_with_totals_tab2 = pd.concat([df_tab2, total_df_tab2], ignore_index=True)

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


# In Tab 3, apply the same sidebar filters to the Tab 3 dataset before rendering it
with tab3:
    st.markdown("""
        <style>
        /* Style the headers for "Select View" and "Select Date Range" */
        .custom-header {
            font-size: 20px;
            font-weight: bold;
            color: #cc0641;  /* Reference color */
            margin-bottom: 5px;  /* Reduced space below header */
            padding: 5px;
            border-left: 5px solid #cc0641; /* Add a colored border on the left */
            background-color: #f9f9f9; /* Light background for contrast */
            border-radius: 3px; /* Slight rounding of edges */
        }

        /* Reduce space between header and radio buttons */
        div[role="radiogroup"] {
            margin-top: -20px;  /* Bring radio buttons closer to the header */
        }
        </style>
        """, unsafe_allow_html=True)

    # Use st.columns to position filters side by side in default Streamlit boxes
    col1, col2 = st.columns(2)

    with col1:
        # Styled header for "Select View"
        st.markdown('<div class="custom-header">Select View:</div>', unsafe_allow_html=True)
        view_type = st.radio("", ['Available Hours', 'Blocked Hours'], key="view_type")

    with col2:
        # Styled header for "Select Date Range"
        st.markdown('<div class="custom-header">Select Date Range:</div>', unsafe_allow_html=True)
        date_range_type = st.radio("", ['Month to Go', 'Total of Month'], key="date_range_type")


    # Define the starting and ending dates for the current month, capping it to today
    today = pd.Timestamp(datetime.now().date())
    start_of_month = today.replace(day=1)
    # Get all the days of the current month up to today, excluding Sundays
    all_days_in_month = pd.date_range(start=start_of_month, end=today)
    all_days_in_month = all_days_in_month[all_days_in_month.weekday != 6]  # Remove Sundays (6 represents Sunday)
    # Filter the columns based on the selected view type
    if view_type == 'Available Hours':
        column_values = 'AvailableHours'
    else:
        column_values = 'BlockedHours'

    # Define color coding for 'Total of Month' or 'Month to Go'
    color_coding_js = JsCode("""
    function(params) {
        var value = params.value;
                             
        if (params.data['Shop[Name]'] === 'Total') {
            return {'font-weight': 'bold', 'backgroundColor': '#e0e0e0'};  // Make the total row bold and set a light background color for visibility
        } else if (value === 0) {
            return {'backgroundColor': '#95cd41', 'color': 'black'};  // Red for 0
        } else if (value > 0 && value < 8) {
            return {'backgroundColor': '#f1b84b', 'color': 'black'};  // Orange for > 0 and < 8
        } else if (value >= 8) {
            return {'backgroundColor': '#cc0641', 'color': 'white'};  // Green for >= 8
        }
        return null;
    }
    """)


    # Custom CSS for the AG-Grid table
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
        ".ag-theme-streamlit .ag-root-wrapper": {
            "border": "2px solid #cc0641",  # Add outer border with specified color
            "border-radius": "5px"  # Optional: Rounded corners for the outer border
        }
    }

    # Prepare the data for "Month to Go" or "Total of Month" by calculating cumulative sums for each day in the month
    pivot_data = []

    if date_range_type == 'Month to Go':
        # Calculate the sum from each day to the end of the month (capped at today)
        for day in all_days_in_month:
            day_data = hcp_data[hcp_data['ShiftDate'] >= day]
            day_sum = day_data.groupby(['Shop[Name]', 'GT_ServiceResource__r.Name'])[column_values].sum().reset_index()
            day_sum.columns = ['Shop Name', 'Resource Name', f'{day.date()} Month to Go']
            pivot_data.append(day_sum)
    else:
        # Calculate the sum from the start of the month to each day (capped at today)
        for day in all_days_in_month:
            day_data = hcp_data[hcp_data['ShiftDate'] <= day]
            day_sum = day_data.groupby(['Shop[Name]', 'GT_ServiceResource__r.Name'])[column_values].sum().reset_index()
            day_sum.columns = ['Shop Name', 'Resource Name', f'{day.date()} Month to Date']
            pivot_data.append(day_sum)

    # Merge the daily data into one DataFrame to display, filling blanks with zeros
    merged_data = pivot_data[0]
    for df in pivot_data[1:]:
        merged_data = pd.merge(merged_data, df, on=['Shop Name', 'Resource Name'], how='outer')

    # Fill in blanks with zeros
    merged_data.fillna(0, inplace=True)

    # Ensure no spaces in field names in the merged DataFrame
    merged_data.columns = [col.replace(' ', '_') for col in merged_data.columns]

    # Dynamic column definitions with conditional formatting
    columnDefs = [
        {
            "headerName": "Shop Name",
            "field": "Shop_Name",
            "resizable": True,
            "flex": 2,  # Adjust flex value to make this column wider
            "minWidth": 150,  # Set a minimum width for columns
            "filter": 'agTextColumnFilter',  # Set filter type to text for shop name
        },
        {
            "headerName": "Resource Name",
            "field": "Resource_Name",
            "resizable": True,
            "flex": 2,
            "minWidth": 150,
            "filter": 'agTextColumnFilter',  # Set filter type to text for resource name
        }
    ]

    # Append dynamic column definitions for each day in the month (excluding Sundays and capped at today)
    for day in all_days_in_month:
        columnDefs.append({
            "headerName": f"{day.date()}",
            "field": f"{day.date()}_{date_range_type.replace(' ', '_')}",
            "valueFormatter": "x.toFixed(1)",  # Format numeric values to 1 decimal place
            "resizable": True,
            "flex": 1,
            "cellStyle": color_coding_js  # Apply color coding based on 'Month to Go' or 'Total of Month' values
        })

    # Calculate totals for numeric columns
    total_row = {'Shop_Name': 'Total', 'Resource_Name': ''}
    for col in merged_data.columns[2:]:
        total_row[col] = merged_data[col].sum()

    # Append the total row to the data
    total_df = pd.DataFrame(total_row, index=[0])
    df_with_totals = pd.concat([merged_data, total_df], ignore_index=True)

    # Configure GridOptionsBuilder with the updated data and column definitions
    gb_tab3 = GridOptionsBuilder.from_dataframe(df_with_totals)

    # Add individual column configurations
    for col_def in columnDefs:
        gb_tab3.configure_column(**col_def)

    # Allow columns to fill the width and use autoHeight for rows
    gb_tab3.configure_grid_options(domLayout='normal', autoSizeColumns='allColumns', enableFillHandle=True)

    # Build grid options
    grid_options_tab3 = gb_tab3.build()

    # Render the AG-Grid in Streamlit with full width and custom styling
    AgGrid(
        df_with_totals,
        gridOptions=grid_options_tab3,
        enable_enterprise_modules=True,
        allow_unsafe_jscode=True,  # Allow JavaScript code execution
        fit_columns_on_grid_load=True,  # Automatically fit columns on load
        height=1000,  # Set grid height to 1000 pixels
        width='100%',  # Set grid width to 100% of the available space
        theme='streamlit',
        custom_css=custom_css  # Apply custom CSS
    )
with tab4:
    # Pivot the table for Tab 4
    pivot_table_tab4 = filtered_hcm.pivot_table(
        index='Shop Name',
        columns='iso_week',
        values=['Duración SF', 'Duración HCM', 'Diferencia de duración'],
        fill_value=0
    )

    # Flatten the columns for display
    pivot_table_tab4.columns = [f"Week {col[1]} {col[0]}" for col in pivot_table_tab4.columns.to_flat_index()]
    pivot_table_tab4_reset = pivot_table_tab4.reset_index()

    # Create the DataFrame for AgGrid
    df_tab4 = pivot_table_tab4_reset

    # Ensure no spaces in field names in df, replacing spaces with underscores or removing them
    df_tab4.columns = [col.replace(' ', '_') for col in df_tab4.columns]

    # Updated JavaScript code to apply color based on Delta value
    js_code = JsCode("""
    function(params) {
        var deltaField = params.colDef.field.replace('Duración_SF', 'Diferencia_de_duración')
                                             .replace('Duración_HCM', 'Diferencia_de_duración');

        var deltaValue = params.data[deltaField];

        if (deltaValue === 0) {
            return {'backgroundColor': '#95cd41', 'color': 'white'};  // Green for Delta = 0
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
            "headerName": "Shop Name",
            "field": "Shop_Name",
            "resizable": True,
            "flex": 2,
            "minWidth": 150
        }
    ]

    # Append dynamic column definitions for each week's SF, HCM, Delta (apply color coding based on Delta value)
    for week in range(35, 41):  # Example for weeks 35 to 40
        columnDefs.append({
            "headerName": f"Week {week}",
            "children": [
                {
                    "field": f"Week_{week}_Duración_SF",
                    "headerName": "SF",
                    "valueFormatter": "x.toFixed(1)",
                    "resizable": True,
                    "flex": 1,
                    "cellStyle": js_code  # Apply color coding based on Delta
                },
                {
                    "field": f"Week_{week}_Duración_HCM",
                    "headerName": "HCM",
                    "valueFormatter": "x.toFixed(1)",
                    "resizable": True,
                    "flex": 1,
                    "cellStyle": js_code  # Apply color coding based on Delta
                },
                {
                    "field": f"Week_{week}_Diferencia_de_duración",
                    "headerName": "Delta",
                    "valueFormatter": "x.toFixed(1)",
                    "resizable": True,
                    "flex": 1,
                    "cellStyle": js_code  # Apply color coding based on Delta
                }
            ]
        })

    # Configure GridOptionsBuilder with JavaScript code and custom styling
    gb_tab4 = GridOptionsBuilder.from_dataframe(df_tab4)

    # Add individual column configurations for conditional formatting
    for column in df_tab4.columns[1:]:
        gb_tab4.configure_column(field=column, cellStyle=js_code)

    # Configure grid options (same as Tab 1)
    gb_tab4.configure_grid_options(domLayout='normal', autoSizeColumns='allColumns', enableFillHandle=True)

    # Build grid options
    grid_options_tab4 = gb_tab4.build()

    # Set the columnDefs in the grid_options dictionary
    grid_options_tab4['columnDefs'] = columnDefs

    # Render the AG-Grid in Streamlit with full width and custom styling
    AgGrid(
        df_tab4,
        gridOptions=grid_options_tab4,
        enable_enterprise_modules=True,
        allow_unsafe_jscode=True,
        fit_columns_on_grid_load=True,
        height=1000,
        width='100%',
        theme='streamlit',
        custom_css=custom_css  # Apply custom CSS
    )
