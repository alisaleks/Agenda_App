import pandas as pd
import pytz
from datetime import datetime, timedelta
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

def load_excel(file_path, usecols=None, **kwargs):
    # Load specific columns if usecols is provided to reduce memory usage
    return pd.read_excel(file_path, usecols=usecols, **kwargs)

def load_csv(file_path, usecols=None, **kwargs):
    # Load specific columns if usecols is provided to reduce memory usage
    return pd.read_csv(file_path, usecols=usecols, **kwargs)

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
'Service Resource[GT_PersonalNumber__c]': str,
'User[GT_StoreCode__c]': str,
'Resource Absence[AbsenceNumber]': str,
'Service Resource[Name]': str
}

# Load datasets with only the necessary columns specified
sfshifts = load_excel(
    'C:/Users/aaleksan/OneDrive - Amplifon S.p.A/Documentos/python_alisa/saturation/Saturation/Satapp/agenda_app/SFshifts_query.xlsx', 
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
    'C:/Users/aaleksan/OneDrive - Amplifon S.p.A/Documentos/python_alisa/saturation/Saturation/Satapp/agenda_app/resource_query.csv',  
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
    'C:/Users/aaleksan/OneDrive - Amplifon S.p.A/Documentos/python_alisa/saturation/Saturation/Satapp/agenda_app/Appointments_aug_oct.xlsx', 
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
    'C:/Users/aaleksan/OneDrive - Amplifon S.p.A/Documentos/python_alisa/saturation/Saturation/Satapp/agenda_app/absences.csv',
    dtype=absences_columns_to_string,
    usecols=[
        'Resource Absence[AbsenceNumber]', 'Resource Absence[Start]', 'Resource Absence[End]', 'Service Resource[Name]', 
        'Service Resource[GT_PersonalNumber__c]', 'User[GT_StoreCode__c]'
    ]
)
# Rename columns to match
absences.rename(columns={
    'Resource Absence[Start]': 'Start',
    'Resource Absence[End]': 'End',
    'Resource Absence[AbsenceNumber]':'AbsenceNumber',
    'Service Resource[Name]': 'Resource.Name',
    'Service Resource[GT_PersonalNumber__c]': 'Resource.GT_PersonalNumber__c', 
    'User[GT_StoreCode__c]': 'Resource.RelatedRecord.GT_StoreCode__c'
}, inplace=True)
# Load regionmapping data
region_mapping_path = 'C:/Users/aaleksan/OneDrive - Amplifon S.p.A/Documentos/python_alisa/saturation/Saturation/Satapp/agenda_app/regionmapping.xlsx'
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
appointments['ApptEndTime'] = pd.to_datetime(appointments['Service Appointment[SchedEndTime]'], errors='coerce').dt.tz_localize(None)
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
# Modify the filtering logic to account for absences that overlap with the start_date and end_date
absences_filtered = absences[(absences['End'] >= start_date) & (absences['Start'] <= end_date)]

# Update the expand_multiday_absences function
def expand_multiday_absences(row):
    start_date = row['Start'].normalize()
    end_date = row['End'].normalize()
    expanded_records = []
    current_date = start_date

    # Define workday hours
    workday_start_hour = 7
    workday_end_hour = 19

    while current_date <= end_date:
        if current_date == start_date and current_date == end_date:
            # Absence starts and ends on the same day
            hours = (row['End'] - row['Start']).total_seconds() / 3600
        elif current_date == start_date:
            # First day: Calculate hours from start time to the end of the workday
            end_of_day = current_date + timedelta(hours=workday_end_hour)
            hours = min((end_of_day - row['Start']).total_seconds() / 3600, workday_end_hour - row['Start'].hour)
        elif current_date == end_date:
            # Last day: Calculate hours from start of workday to end time
            start_of_day = current_date + timedelta(hours=workday_start_hour)
            hours = min((row['End'] - start_of_day).total_seconds() / 3600, row['End'].hour - workday_start_hour)
        else:
            # Full day: Full workday hours for days entirely within the absence period
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
            'Resource.Name': row['Resource.Name']
        })
        current_date += timedelta(days=1)
    
    return expanded_records

# Filter absences to include any absence overlapping with the period
expanded_absences = absences_filtered.apply(expand_multiday_absences, axis=1)
expanded_absences = pd.DataFrame([record for sublist in expanded_absences for record in sublist])

# Group expanded absences by PersonalNumberKey and AbsenceDate
absences_grouped = expanded_absences.groupby(['PersonalNumberKey', 'AbsenceDate']).agg({
    'AbsenceDurationHours': 'sum',  # Sum of absence duration hours
    'Resource.GT_PersonalNumber__c': 'first', 
    'Resource.RelatedRecord.GT_StoreCode__c': 'first',  
    'AbsenceNumber': 'first',  
    'Resource.Name': 'first',  
    'AbsenceStartTime': 'first',
    'AbsenceEndTime': 'last'
}).reset_index()


absences_grouped[absences_grouped['PersonalNumberKey'] == '969_25367']
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

check= sfshifts_merged[sfshifts_merged['PersonalNumberKey'] == '969_25367']
check
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
a=shift_slots[(shift_slots['GT_ShopCode__c']=='969') &
    (shift_slots['date'] >= pd.Timestamp('2024-09-05')) &
    (shift_slots['date'] < pd.Timestamp('2024-09-06'))]

print(a[['TotalSlots_gross', 'TotalSlots', 'OpenHours', 'ShiftDurationMinutesAdjusted', 'BookedHours',  'Shop[Name]', 'date']])
shift_slots.columns
# Define the filename and path for the Excel file
output_file_path = 'shiftslots.xlsx'

# Export the DataFrame to Excel
shift_slots.to_excel(output_file_path, index=False, engine='openpyxl')  # You can also use 'xlsxwriter' if you prefer