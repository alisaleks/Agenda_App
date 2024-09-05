
import sys
import streamlit as st
import pandas as pd
import pytz
from datetime import datetime, timedelta
from st_aggrid import GridOptionsBuilder, AgGrid, JsCode
from st_aggrid.shared import GridUpdateMode, DataReturnMode, ColumnsAutoSizeMode, AgGridTheme, ExcelExportMode
from st_aggrid.AgGridReturn import AgGridReturn
import json

# Function to handle out-of-bound datetime values
def handle_out_of_bound_dates(date_str):
    try:
        return pd.to_datetime(date_str)
    except (pd.errors.OutOfBoundsDatetime, OverflowError):
        return pd.Timestamp.max

# Function to check if a resource is active
def is_active(row, start_date, end_date):
    if pd.isnull(row['EffectiveEndDate']) or pd.isnull(row['EffectiveStartDate']):
        return False
    return not (row['EffectiveEndDate'] < start_date or row['EffectiveStartDate'] > end_date)

utc = pytz.timezone('UTC')
madrid = pytz.timezone('Europe/Madrid')
def convert_utc_to_madrid(dt):
    if pd.isnull(dt):
        return dt
    dt_utc = utc.localize(dt)
    dt_madrid = dt_utc.astimezone(madrid)
    return dt_madrid

@st.cache_data
def load_excel(file_path, usecols=None, **kwargs):
    # Load specific columns if usecols is provided to reduce memory usage
    return pd.read_excel(file_path, usecols=usecols, **kwargs)

@st.cache_data
def load_csv(file_path, usecols=None, **kwargs):
    # Load specific columns if usecols is provided to reduce memory usage
    return pd.read_csv(file_path, usecols=usecols, **kwargs)


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

shifts_columns_to_string = {
'Shift[ShiftNumber]': str,
'Shift[Label]': str,
'Service Resource[Name]': str,
'Shop[GT_ShopCode__c]': str,
'Service Resource[GT_Role__c]': str,
'Shift[StartTime]': str,
'Shift[EndTime]': str,
'Shift[ServiceResourceId]': str,
'Shop[GT_CountryCode__c]': str,
'Shop[Country]': str,
'Shop[Name]': str,
'Shop[GT_AreaManagerCode__c]': str,
'Shift[LastModifiedDate]': str,
'Service Resource[GT_PersonalNumber__c]': str,
'Shop[GT_StoreType__c]': str
}
resources_columns_to_string = {
'Shop[GT_CountryCode__c]': str,
'Shop[Country]': str,
'Service Territory Member[ServiceTerritoryId]': str,
'Shop[GT_ShopCode__c]': str,
'Service Resource[Name]': str,
'Service Territory Member[ServiceResourceId]': str,
'Service Resource[Name].1': str,
'Service Territory Member[EffectiveStartDate]': str,
'Service Territory Member[EffectiveEndDate]': str,
'Service Resource[GT_Role__c]': str,
'Service Resource[GT_PersonalNumber__c]': str
}
appointments_columns_to_string = {
'Service Appointment[AppointmentNumber]': str,
'Service Appointment[ServiceTerritoryId]': str,
'Service Appointment[Business_Shop__c]': str,
'Service Appointment[GT_ShopCode__c]': str,
'Shop[GT_CountryCode__c]': str,
'Service Appointment[GT_Cluster__c]': str,
'Service Appointment[GT_Macrocategory__c]': str,
'Service Appointment[GT_AccountNameConcatenated__c]': str,
'Shop[GT_AreaCode__c]': str,
'Shop[GT_StoreType__c]': str,
'Shop[GT_AreaManagerCode__c]': str,
'Service Appointment[SchedStartTime]': str,
'Service Appointment[SchedEndTime]': str,
'Service Resource[GT_Role__c]': str,
'Service Appointment[GT_ServiceResource__c]': str,
'Service Resource[Name]': str,
'Service Appointment[Status]': str,
'Service Appointment[LastModifiedDate]': str
}
absences_columns_to_string = {
'Resource.GT_PersonalNumber__c': str,
'Resource.RelatedRecord.GT_StoreCode__c': str,
'Resource.Account': str,
'AbsenceNumber': str,
'Resource.Name': str
}
# Load datasets with specific column types
# Load datasets with only the necessary columns specified
sfshifts = load_excel( 'SFshifts_query.xlsx', 
    dtype=shifts_columns_to_string,
    usecols=[
        'Shift[ShiftNumber]', 'Shift[Label]', 'Service Resource[Name]', 'Shop[GT_ShopCode__c]', 
        'Service Resource[GT_Role__c]', 'Shift[StartTime]', 'Shift[EndTime]', 
        'Shift[ServiceResourceId]', 'Shop[GT_CountryCode__c]', 'Shop[Country]', 
        'Shop[Name]', 'Shop[GT_AreaManagerCode__c]', 'Shift[LastModifiedDate]', 
        'Service Resource[GT_PersonalNumber__c]', 'Shop[GT_StoreType__c]', 'Shop[GT_AreaCode__c]'
    ]
)

resources = load_csv(
   'resource_query.csv',  
    dtype=resources_columns_to_string,
    usecols=[
        'Shop[GT_CountryCode__c]', 'Service Territory Member[EffectiveEndDate]', 
        'Service Territory Member[EffectiveStartDate]', 'Shop[Country]', 
        'Service Territory Member[ServiceTerritoryId]', 'Shop[GT_ShopCode__c]', 
        'Service Territory Member[ServiceResourceId]', 'Service Resource[GT_PersonalNumber__c]', 
        'Service Resource[GT_Role__c]'
    ]
)

appointments = load_excel(
    'Appointments_aug_oct.xlsx', 
    dtype=appointments_columns_to_string,
    usecols=[
        'Service Appointment[AppointmentNumber]', 'Service Appointment[ServiceTerritoryId]', 
        'Service Appointment[Business_Shop__c]', 'Service Appointment[GT_ShopCode__c]', 
        'Shop[GT_CountryCode__c]', 'Service Appointment[GT_Cluster__c]', 
        'Service Appointment[GT_Macrocategory__c]', 'Service Appointment[GT_AccountNameConcatenated__c]', 
        'Shop[GT_AreaCode__c]', 'Shop[GT_StoreType__c]', 'Shop[GT_AreaManagerCode__c]', 
        'Service Appointment[SchedStartTime]', 'Service Appointment[SchedEndTime]', 
        'Service Resource[GT_Role__c]', 'Service Appointment[GT_ServiceResource__c]', 
        'Service Resource[Name]', 'Service Appointment[Status]', 'Service Appointment[LastModifiedDate]'
    ]
)

absences = load_csv(
    'absences.csv',
    dtype=absences_columns_to_string,
    usecols=[
        'AbsenceNumber', 'Start', 'End', 'Resource.Name', 'Resource.Account', 
        'Resource.GT_PersonalNumber__c', 'Resource.RelatedRecord.GT_StoreCode__c', 'Type'
    ]
)
# Load regionmapping data
region_mapping_path = 'regionmapping.xlsx'
region_mapping = load_excel(region_mapping_path)

sfshifts['StartTime'] = pd.to_datetime(sfshifts['Shift[StartTime]'], errors='coerce')
sfshifts['EndTime'] = pd.to_datetime(sfshifts['Shift[EndTime]'], errors='coerce')
start_date = datetime(2024, 9, 1)
end_date = datetime(2024, 10, 6)
shifts_filtered = sfshifts[(sfshifts['StartTime'] >= start_date) & (sfshifts['EndTime'] <= end_date)].copy()

# Rename columns to match
shifts_filtered.rename(columns={
    'Shop[GT_ShopCode__c]': 'GT_ShopCode__c',
    'Service Resource[Name]': 'GT_ServiceResource__r.Name'
}, inplace=True)

resources.rename(columns={
    'Shop[GT_ShopCode__c]': 'GT_ShopCode__c'
}, inplace=True)

# Convert specific columns to datetime
shifts_filtered['LastModifiedDate'] = pd.to_datetime(shifts_filtered['Shift[LastModifiedDate]'], errors='coerce')
appointments['ApptStartTime'] = pd.to_datetime(appointments['Service Appointment[SchedStartTime]'], errors='coerce').dt.tz_localize(None)
appointments['ApptEndTime'] = pd.to_datetime(
    appointments['Service Appointment[SchedEndTime]'],
    errors='coerce'
).dt.tz_localize(None)
failed_parsing = appointments[appointments['ApptEndTime'].isna()]
appointments.loc[failed_parsing.index, 'Service Appointment[SchedEndTime]'] = '2024-08-27 9:35:00' 
appointments['ApptEndTime'] = pd.to_datetime(
    appointments['Service Appointment[SchedEndTime]'],
    errors='coerce'
).dt.tz_localize(None)
appointments['ApptsLastModifiedDate'] = pd.to_datetime(appointments['Service Appointment[LastModifiedDate]'], errors='coerce').dt.tz_localize(None)
# Drop original datetime columns
shifts_filtered.drop(columns=['Shift[StartTime]', 'Shift[EndTime]', 'Shift[LastModifiedDate]'], inplace=True)
appointments.drop(columns=['Service Appointment[SchedStartTime]', 'Service Appointment[SchedEndTime]', 'Service Appointment[LastModifiedDate]'], inplace=True)
# Convert to datetime with out-of-bound handling for specific columns
resources['EffectiveEndDate'] = resources['Service Territory Member[EffectiveEndDate]'].apply(handle_out_of_bound_dates)
resources['EffectiveStartDate'] = resources['Service Territory Member[EffectiveStartDate]'].apply(handle_out_of_bound_dates)

# Add a date column
shifts_filtered['date'] = shifts_filtered['StartTime'].dt.strftime('%d/%m/%Y')

# Directly extract ISO week and year from StartTime
shifts_filtered['iso_week'] = shifts_filtered['StartTime'].dt.isocalendar().week
shifts_filtered['iso_year'] = shifts_filtered['StartTime'].dt.isocalendar().year

shifts_filtered['StartDateHour'] = shifts_filtered['StartTime'].dt.strftime('%Y-%m-%d %H:00:00')
shifts_filtered['Key'] = shifts_filtered['GT_ShopCode__c'] + '_' + shifts_filtered['GT_ServiceResource__r.Name'] + '_' + shifts_filtered['StartDateHour']
duplicates = shifts_filtered[shifts_filtered.duplicated(subset=['Key'], keep=False)]
shifts_filtered = shifts_filtered.sort_values(by=['Key', 'LastModifiedDate'], ascending=[True, False])
shifts_filtered = shifts_filtered.drop_duplicates(subset=['Key'], keep='first')
shifts_filtered['ShopResourceKey'] = shifts_filtered['GT_ShopCode__c'] + shifts_filtered['Shift[ServiceResourceId]']
resources['ShopResourceKey'] = resources['GT_ShopCode__c'] + resources['Service Territory Member[ServiceResourceId]']

resources['IsActive'] = resources.apply(is_active, axis=1, args=(start_date, end_date))
active_resources = resources[resources['IsActive']]
shifts_filtered = shifts_filtered[shifts_filtered['ShopResourceKey'].isin(active_resources['ShopResourceKey'])]

shifts_filtered['PersonalNumberKey'] = shifts_filtered['GT_ShopCode__c'] + '_' + shifts_filtered['Service Resource[GT_PersonalNumber__c]']
shifts_filtered['ShiftDate'] = shifts_filtered['StartTime'].dt.date
shifts_filtered['ShiftDurationHours'] = (shifts_filtered['EndTime'] - shifts_filtered['StartTime']).dt.total_seconds() / 3600

absences['Start'] = pd.to_datetime(absences['Start'], errors='coerce')
absences['End'] = pd.to_datetime(absences['End'], errors='coerce')
absences['AbsenceDurationHours'] = (absences['End'] - absences['Start']).dt.total_seconds() / 3600
absences['PersonalNumberKey'] = absences['Resource.RelatedRecord.GT_StoreCode__c'] + '_' + absences['Resource.GT_PersonalNumber__c']

# Group absences by PersonalNumberKey and date to find total absence hours per day per resource
absences['AbsenceDate'] = absences['Start'].dt.date
def expand_multiday_absences(row):
    start_date = row['Start'].normalize()
    end_date = row['End'].normalize()
    expanded_records = []
    current_date = start_date

    # Assuming a standard 12-hour workday from 9:00 to 17:00
    workday_start_hour = 7
    workday_end_hour = 19

    while current_date <= end_date:
        if current_date == start_date and current_date == end_date:
            # Absence starts and ends on the same day
            hours = (row['End'] - row['Start']).total_seconds() / 3600
        elif current_date == start_date:
            # First day: Calculate hours from start time to the end of the workday
            end_of_day = current_date + timedelta(hours=workday_end_hour)
            hours = ((end_of_day - row['Start']).total_seconds() / 3600)
        elif current_date == end_date:
            # Last day: Calculate hours from start of workday to end time
            start_of_day = current_date + timedelta(hours=workday_start_hour)
            hours = ((row['End'] - start_of_day).total_seconds() / 3600)
        else:
            # Full day: 8 hours of absence for a full day between start and end dates
            hours = workday_end_hour - workday_start_hour

        expanded_records.append({
            'PersonalNumberKey': row['PersonalNumberKey'],
            'AbsenceDate': current_date.date(),
            'AbsenceDurationHours': hours,
            'AbsenceStartTime': row['Start'],
            'AbsenceEndTime': row['End'],
            'AbsenceNumber': row['AbsenceNumber'],
            'Resource.GT_PersonalNumber__c': row['Resource.GT_PersonalNumber__c'],
            'Resource.RelatedRecord.GT_StoreCode__c': row['Resource.RelatedRecord.GT_StoreCode__c'],
            'Resource.Account': row['Resource.Account'],
            'Resource.Name': row['Resource.Name'],
            'Type': row['Type'], 
        })
        current_date += timedelta(days=1)
    
    return expanded_records

# Expand all absences into daily records
expanded_absences = absences.apply(expand_multiday_absences, axis=1)
expanded_absences = pd.DataFrame([record for sublist in expanded_absences for record in sublist])
absences_grouped = expanded_absences.groupby(['PersonalNumberKey', 'AbsenceDate']).agg({
    'AbsenceDurationHours': 'sum',  # Sum of absence duration hours
    'Resource.GT_PersonalNumber__c': 'first', 
    'Resource.RelatedRecord.GT_StoreCode__c': 'first',  
    'Resource.Account': 'first',  
    'AbsenceNumber': 'first',  
    'Resource.Name': 'first',  
    'Type'  : 'first',
    'AbsenceStartTime' : 'first',
    'AbsenceEndTime' : 'last'
}).reset_index()


# Group shifts by PersonalNumberKey and ShiftDate to find total shift hours per day per resource
shifts_grouped = shifts_filtered.groupby(['PersonalNumberKey', 'ShiftDate']).agg({
    'ShiftDurationHours': 'sum',  # Sum of absence duration hours
    'Service Resource[GT_Role__c]' : 'first', 
    'GT_ServiceResource__r.Name' : 'first',
    'GT_ShopCode__c': 'first',
    'Shift[Label]': 'first',
    'ShopResourceKey': 'first',  
    'StartDateHour': 'first',  
    'iso_year': 'first',  
    'iso_week': 'first',
    'date': 'first',
    'LastModifiedDate' : 'first',
    'GT_ShopCode__c': 'first',
    'Shop[Name]': 'first',
    'Shop[GT_CountryCode__c]': 'first',
    'Shop[Country]': 'first',
    'Shop[GT_AreaManagerCode__c]': 'first',
    'Shop[GT_AreaCode__c]' : 'first',
    'Shop[GT_StoreType__c]': 'first',
    'StartTime': 'first',
    'EndTime': 'last'
    
}).reset_index()

# Merge expanded_absences with shifts data to calculate adjusted shift hours
sfshifts_merged = pd.merge(
    shifts_grouped,
    absences_grouped,
    how='left',
    left_on=['PersonalNumberKey', 'ShiftDate'],
    right_on=['PersonalNumberKey','AbsenceDate'],
    suffixes=('', '_absence')
)

# Calculate the adjusted shift duration by subtracting the absence hours
sfshifts_merged['ShiftDurationHoursAdjusted'] = sfshifts_merged['ShiftDurationHours'] - sfshifts_merged['AbsenceDurationHours'].fillna(0)
# replacing negatives to 0
sfshifts_merged['ShiftDurationHoursAdjusted'] = sfshifts_merged['ShiftDurationHoursAdjusted'].apply(lambda x: max(x, 0))
# Recalculate ShiftDurationMinutes based on adjusted hours
sfshifts_merged['ShiftDurationMinutesAdjusted'] = sfshifts_merged['ShiftDurationHoursAdjusted'] * 60
sfshifts_merged['ShiftDate'] = pd.to_datetime(sfshifts_merged['ShiftDate'])

filtered_df = sfshifts_merged[
    sfshifts_merged['PersonalNumberKey'].str.startswith('049') & 
    (sfshifts_merged['ShiftDate'] == '2024-09-02') ]           
# Convert the Hora de inicio and Hora de fin columns to Madrid time
sfshifts_merged['Hora de inicio'] = sfshifts_merged['StartTime'].apply(convert_utc_to_madrid)
sfshifts_merged['Hora de fin'] = sfshifts_merged['EndTime'].apply(convert_utc_to_madrid)
appointments['Horario de inicio de la cita'] = appointments['ApptStartTime'].apply(convert_utc_to_madrid)
appointments['Horario de fin de la cita'] = appointments['ApptEndTime'].apply(convert_utc_to_madrid)
# Check data for a specific shop and date before the pivot

# Remove the timezone information (make them timezone-naive)
sfshifts_merged['Hora de inicio'] = sfshifts_merged['Hora de inicio'].dt.tz_localize(None)
sfshifts_merged['Hora de fin'] = sfshifts_merged['Hora de fin'].dt.tz_localize(None)
appointments['Horario de inicio de la cita'] = appointments['Horario de inicio de la cita'].dt.tz_localize(None)
appointments['Horario de fin de la cita'] = appointments['Horario de fin de la cita'].dt.tz_localize(None)

check= sfshifts_merged[(sfshifts_merged['PersonalNumberKey'] == '005_4375') &
                                     (sfshifts_merged['ShiftDate'] >= pd.to_datetime('2024-09-01')) &
                                     (sfshifts_merged['ShiftDate'] < pd.to_datetime('2024-09-04'))]
# Identify duplicates based on the specified subset of columns
duplicate_mask = appointments.duplicated(subset=[
    'Shop[GT_CountryCode__c]',
    'Service Appointment[GT_ShopCode__c]',
    'Service Resource[Name]',
    'Service Appointment[GT_AccountNameConcatenated__c]',
    'ApptStartTime',
    'ApptEndTime'
], keep=False)

# Filter the DataFrame to show only duplicates
duplicates = appointments[duplicate_mask]

# Sort by the specified subset of columns and 'ApptsLastModifiedDate'
appointments = appointments.sort_values(by=[
    'Shop[GT_CountryCode__c]',
    'Service Appointment[GT_ShopCode__c]',
    'Service Resource[Name]',
    'Service Appointment[GT_AccountNameConcatenated__c]',
    'ApptStartTime',
    'ApptEndTime',
    'ApptsLastModifiedDate'
], ascending=[True, True, True, True, True, True, False])

# Drop duplicates, keeping only the last modified
appointments = appointments.drop_duplicates(subset=[
    'Shop[GT_CountryCode__c]',
    'Service Appointment[GT_ShopCode__c]',
    'Service Resource[Name]',
    'Service Appointment[GT_AccountNameConcatenated__c]',
    'ApptStartTime',
    'ApptEndTime'
], keep='first')
# Calculate the total number of 5-minute slots available per shop and date
shift_slots = sfshifts_merged.groupby(['GT_ShopCode__c', 'Shop[Name]', 'date'])[['ShiftDurationMinutesAdjusted', 'ShiftDurationHours']].sum().reset_index()
shift_slots['TotalSlots'] = shift_slots['ShiftDurationMinutesAdjusted'] / 5
shift_slots['TotalSlots_gross'] = shift_slots['ShiftDurationHours']*60 / 5


# Filter appointments within August
appointments_filtered = appointments[(appointments['ApptStartTime'] >= start_date) & (appointments['ApptEndTime'] <= end_date)].copy()

# Example: Deleting unused DataFrames to free up memory
del sfshifts, resources, appointments, absences

# Fill NA in categories with 'First Visit'
appointments_filtered['Service Appointment[GT_Macrocategory__c]'] = appointments_filtered['Service Appointment[GT_Macrocategory__c]'].fillna('First Visit')
appointments_filtered['Service Appointment[GT_Macrocategory__c]'] = appointments_filtered['Service Appointment[GT_Macrocategory__c]'].str.strip()
appointments_filtered['Service Appointment[GT_Macrocategory__c]'] = appointments_filtered['Service Appointment[GT_Macrocategory__c]'].replace({
    'Fitting': 'Pre-Sales',
    'Post-Sales': 'After-Sales'
})


# Step 1: Generate all 5-minute slots for all appointments across all categories
slots = []

# Step 1: Generate all 5-minute slots for all appointments across all categories
for _, row in appointments_filtered.iterrows():
    slot_start = row['ApptStartTime']
    slot_end = row['ApptEndTime']
    
    # Generate all 5-minute slots for this appointment
    while slot_start < slot_end:
        slots.append((row['Service Appointment[GT_ShopCode__c]'], row['Service Resource[Name]'], slot_start.strftime('%d/%m/%Y'), slot_start, row['ApptsLastModifiedDate']))
        slot_start += timedelta(minutes=5)

# Create a DataFrame from all the slots including the last modified date for duplicate removal
all_slots_df = pd.DataFrame(slots, columns=['GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot', 'LastModifiedDate'])

# Ensure the 'date' column is of datetime type for proper merging
all_slots_df['date'] = pd.to_datetime(all_slots_df['date'], format='%d/%m/%Y')

# Sort slots by 'GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot', and 'LastModifiedDate' to prioritize keeping the earliest slot
all_slots_df = all_slots_df.sort_values(by=['GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot', 'LastModifiedDate'], ascending=[True, True, True, True, False])

# Remove duplicate 5-minute slots, keeping only the earliest modified slot
all_slots_df = all_slots_df.drop_duplicates(subset=['GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot'], keep='first')

# Step 2: Count the unique slots, ensuring no double-counting for overlaps
net_booked_slots = all_slots_df.groupby(['GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot']).size().reset_index(name='Count')

# Normalize counts to 1 for any slot count greater than 1
net_booked_slots['Count'] = net_booked_slots['Count'].apply(lambda x: 1 if x > 0 else 0)

# Calculate the total booked slots by shop, service resource, and date
total_booked_slots_by_date = net_booked_slots.groupby(['GT_ShopCode__c', 'Service Resource[Name]', 'date'])['Count'].sum().reset_index()

# Step 4: Recalculate `TotalBookedSlots` based on the total booked slots by date
shift_slots['date'] = pd.to_datetime(shift_slots['date'], format='%d/%m/%Y', errors='coerce')
total_booked_slots_by_date['date'] = pd.to_datetime(total_booked_slots_by_date['date'], errors='coerce')

grouped_df = total_booked_slots_by_date.groupby(['GT_ShopCode__c', 'date']).agg({'Count': 'sum'}).reset_index()
shift_slots = pd.merge(
    shift_slots, 
    grouped_df, 
    on=['GT_ShopCode__c', 'date'], 
    how='left'
)
shift_slots['TotalBookedSlots'] = shift_slots['Count'].fillna(0)
shift_slots.drop(columns=['Count'], inplace=True)

shift_slots['OpenSlots'] = shift_slots['TotalSlots'] - shift_slots['TotalBookedSlots']
shift_slots['OpenSlots'] = shift_slots['OpenSlots'].apply(lambda x: max(x, 0))

# Check for rows where OpenSlots was negative (if any remain)
negative_open_slots = shift_slots[shift_slots['OpenSlots'] < 0]
shift_slots['OpenHours'] = (shift_slots['OpenSlots'] * 5) / 60
shift_slots['BookedHours'] = (shift_slots['TotalBookedSlots'] * 5) / 60
shift_slots['TotalHours'] = (shift_slots['TotalSlots_gross'] * 5) / 60
shift_slots['SaturationPercentage'] = (shift_slots['BookedHours'] / shift_slots['TotalHours']) * 100
shift_slots['SaturationPercentage'] = shift_slots['SaturationPercentage'].clip(lower=0, upper=100)

# Add weekday name and ISO week
shift_slots['date'] = pd.to_datetime(shift_slots['date'], format='%d/%m/%Y')
shift_slots['day'] = shift_slots['date'].dt.day
shift_slots['weekday'] = shift_slots['date'].dt.day_name()
shift_slots['iso_week'] = shift_slots['date'].dt.isocalendar().week
shift_slots['month'] = shift_slots['date'].dt.strftime('%B')
# Remove Sundays from the shift_slots DataFrame
shift_slots = shift_slots[shift_slots['weekday'] != 'Sunday']
check= shift_slots[shift_slots['GT_ShopCode__c'] == '240']
# Sort the shift_slots DataFrame by 'date' to ensure the correct order of days
shift_slots = shift_slots.sort_values(by='date')
# Merge cleaned_data with region_mapping to update 'Region_Descr'
shift_slots = shift_slots.merge(region_mapping[['CODE', 'REGION', 'AREA']], left_on='GT_ShopCode__c', right_on='CODE', how='left')
shift_slots.rename(columns={'REGION': 'Region', 'AREA': 'Area'}, inplace=True)
shift_slots = shift_slots.drop(columns=['CODE'])

# Sidebar filters (all converted to single-selection using selectbox)
iso_week_filter = st.sidebar.selectbox('Select ISO Week', sorted(shift_slots['iso_week'].unique()))

# Calculate the previous ISO week and year based on the selected ISO week
selected_iso_year = datetime.now().year  # Assuming current year, adjust if you have a different dataset
previous_iso_week = iso_week_filter - 1
previous_iso_year = selected_iso_year
# Handle ISO year transition if the selected week is the first week of the year
if previous_iso_week == 0:
    previous_iso_year -= 1
    previous_iso_week = 52 if (pd.Timestamp(f"{previous_iso_year}-12-28").isocalendar()[1] == 52) else 53


# Sidebar filters for Region and Area
region_list = sorted(shift_slots['Region'].dropna().unique().tolist())
region_options = ["All"] + region_list

selected_region = st.sidebar.selectbox(
    'Select Region:',
    options=region_options,
    index=0,  # Default to "All"
    help="Select a region or 'All' to view data for all regions."
)

area_list = sorted(shift_slots['Area'].dropna().unique().tolist())
area_options = ["All"] + area_list

selected_area = st.sidebar.selectbox(
    'Select Area:',
    options=area_options,
    index=0,  # Default to "All"
    help="Select an area or 'All' to view data for all areas."
)

# Initialize shop filter with "All" as an option
shop_list = sorted(shift_slots['Shop[Name]'].unique().tolist())
shop_options = ["All"] + shop_list

# Single-select dropdown for shop filter with "All" as an option
selected_shop = st.sidebar.selectbox(
    'Select Shop:',
    options=shop_options,
    index=0,  # Default to "All"
    help="Select a shop or 'All' to view data for all shops."
)

# Function for lazy loading of filtered data
@st.cache_data
def filter_shift_slots(data, iso_week_filter, selected_shop, selected_region, selected_area):
    # Start with the full data
    filtered_data = data.copy()

    # Apply filters based on sidebar selections
    filtered_data = filtered_data[filtered_data['iso_week'] == iso_week_filter]

    if selected_region != "All":
        filtered_data = filtered_data[filtered_data['Region'] == selected_region]

    if selected_area != "All":
        filtered_data = filtered_data[filtered_data['Area'] == selected_area]
        
    if selected_shop != "All":
        filtered_data = filtered_data[filtered_data['Shop[Name]'] == selected_shop]

    return filtered_data

# Filter data for the selected ISO week using the user's filter choices
filtered_data = filter_shift_slots(shift_slots, iso_week_filter, selected_shop, selected_region, selected_area)

# Check if filtered data is empty after applying the filters
if filtered_data.empty:
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
selected_week_start = pd.Timestamp(selected_iso_year, 1, 1) + pd.offsets.Week(weekday=0) * (iso_week_filter - 1)
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
numeric_cols = ['OpenHours', 'TotalHours', 'SaturationPercentage']
filtered_data_numeric = filtered_data[numeric_cols + ['day', 'weekday', 'GT_ShopCode__c', 'Shop[Name]']]

# Aggregating data by GT_ShopCode__c, Shop[Name], date, and weekday
aggregated_data = filtered_data.groupby(['GT_ShopCode__c', 'Shop[Name]', 'date', 'weekday']).agg(
    OpenHours=('OpenHours', 'sum'),
    TotalHours=('TotalHours', 'sum'),
    SaturationPercentage=('SaturationPercentage', 'mean')
).reset_index()

aggregated_data['date'] = pd.to_datetime(aggregated_data['date']).dt.date
check= aggregated_data[(aggregated_data['GT_ShopCode__c'] == '240')]

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

    if (totalHoursValue === 0) {
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