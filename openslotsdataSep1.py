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
'Service Resource[Name]': str,
'Shop[GT_ShopCode__c]': str,
'Shift[StartTime]': str,
'Shift[EndTime]': str,
'Shift[ServiceResourceId]': str,
'Shop[Name]': str,
'Service Resource[GT_PersonalNumber__c]': str,
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
'Service Resource[IsActive]' : str,
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
'Service Resource[Id]':str
}

# Load datasets with only the necessary columns specified
sfshifts = load_excel(
    'C:/Users/aaleksan/OneDrive - Amplifon S.p.A/Documentos/python_alisa/saturation/Saturation/Satapp/agenda_app/SFshifts_query_1Sep.xlsx', 
    dtype=shifts_columns_to_string,
    usecols=['Service Resource[Name]', 'Shop[GT_ShopCode__c]', 'Shift[StartTime]', 'Shift[EndTime]', 
        'Shift[ServiceResourceId]','Shop[Name]',
        'Service Resource[GT_PersonalNumber__c]'
    ]
)
sfshifts.head()
resources = load_csv(
    'C:/Users/aaleksan/OneDrive - Amplifon S.p.A/Documentos/python_alisa/saturation/Saturation/Satapp/agenda_app/resource_query.csv',  
    dtype=resources_columns_to_string,
    usecols=[
        'Shop[GT_CountryCode__c]', 'Service Territory Member[EffectiveEndDate]', 
        'Service Territory Member[EffectiveStartDate]', 'Shop[Country]', 
        'Service Territory Member[ServiceTerritoryId]', 'Shop[GT_ShopCode__c]', 
        'Service Territory Member[ServiceResourceId]', 'Service Resource[GT_PersonalNumber__c]', 'Service Resource[IsActive]',
        'Service Resource[GT_Role__c]'
    ], 
)

absences = load_excel(
    'C:/Users/aaleksan/OneDrive - Amplifon S.p.A/Documentos/python_alisa/saturation/Saturation/Satapp/agenda_app/absences_sep1.xlsx',
    dtype=absences_columns_to_string,
    usecols=['Resource Absence[Start]', 'Resource Absence[End]','Service Resource[GT_PersonalNumber__c]', 'User[GT_StoreCode__c]', 'Service Resource[Id]'
    ]
)
# Rename columns to match
absences.rename(columns={
    'Resource Absence[Start]': 'Start',
    'Resource Absence[End]': 'End',
    'Service Resource[GT_PersonalNumber__c]': 'Resource.GT_PersonalNumber__c', 
    'User[GT_StoreCode__c]': 'Resource.RelatedRecord.GT_StoreCode__c'
}, inplace=True)
# Load regionmapping data
region_mapping_path = 'C:/Users/aaleksan/OneDrive - Amplifon S.p.A/Documentos/python_alisa/saturation/Saturation/Satapp/agenda_app/regionmapping.xlsx'
region_mapping = load_excel(region_mapping_path)
region_mapping.columns
# Filter the original region_mapping DataFrame
region_mapping = region_mapping[region_mapping['SYM'] != 'N']
sfshifts.head()
sfshifts['StartTime'] = pd.to_datetime(sfshifts['Shift[StartTime]'], errors='coerce')
sfshifts['EndTime'] = pd.to_datetime(sfshifts['Shift[EndTime]'], errors='coerce')
sfshifts['StartTime'] = sfshifts['StartTime'].dt.tz_convert(None)
sfshifts['EndTime'] = sfshifts['EndTime'].dt.tz_convert(None)
start_date = datetime(2024, 9, 2) 
end_date = datetime(2024, 10, 6)
shifts_filtered = sfshifts[(sfshifts['StartTime'] >= start_date) & (sfshifts['EndTime']<= end_date)].copy()
# Rename columns to match
shifts_filtered.rename(columns={
    'Shop[GT_ShopCode__c]': 'GT_ShopCode__c',
    'Service Resource[Name]': 'GT_ServiceResource__r.Name'
}, inplace=True)
print(shifts_filtered[shifts_filtered['GT_ShopCode__c'] == '978'])

resources.rename(columns={
    'Shop[GT_ShopCode__c]': 'GT_ShopCode__c'
}, inplace=True)
# Convert specific columns to datetime
# Drop original datetime columns
shifts_filtered.drop(columns=['Shift[StartTime]', 'Shift[EndTime]'], inplace=True)
shifts_filtered['PersonalNumberKey'] = shifts_filtered['GT_ShopCode__c'] + '_' + shifts_filtered['Service Resource[GT_PersonalNumber__c]']

# Convert to datetime with out-of-bound handling for specific columns
resources['EffectiveEndDate'] = resources['Service Territory Member[EffectiveEndDate]'].apply(handle_out_of_bound_dates)
resources['EffectiveStartDate'] = resources['Service Territory Member[EffectiveStartDate]'].apply(handle_out_of_bound_dates)

# Add a date column
shifts_filtered['date'] = shifts_filtered['StartTime'].dt.strftime('%d/%m/%Y')
shifts_filtered['ShiftDate'] = shifts_filtered['StartTime'].dt.date
# Directly extract ISO week and year from StartTime
shifts_filtered['iso_week'] = shifts_filtered['StartTime'].dt.isocalendar().week
shifts_filtered['iso_year'] = shifts_filtered['StartTime'].dt.isocalendar().year

shifts_filtered['StartDateHour'] = shifts_filtered['StartTime'].dt.strftime('%Y-%m-%d %H:00:00')
shifts_filtered['Key'] = shifts_filtered['GT_ShopCode__c'] + '_' + shifts_filtered['GT_ServiceResource__r.Name'] + '_' + shifts_filtered['StartDateHour']
#duplicate treatment
duplicates = shifts_filtered[shifts_filtered.duplicated(subset=['Key'], keep=False)]
shifts_filtered = shifts_filtered.drop_duplicates(subset=['Key'], keep='first')

shifts_filtered['ShiftDurationHours'] = (shifts_filtered['EndTime'] - shifts_filtered['StartTime']).dt.total_seconds() / 3600

shifts_filtered['ShopResourceKey'] = shifts_filtered['GT_ShopCode__c'] + shifts_filtered['Shift[ServiceResourceId]']
resources['ShopResourceKey'] = resources['GT_ShopCode__c'] + resources['Service Territory Member[ServiceResourceId]']

check = shifts_filtered[(shifts_filtered['PersonalNumberKey'] == '969_25367') & (pd.to_datetime(shifts_filtered['ShiftDate'], errors='coerce') == '2024-09-02')]
check
# Step 1: Convert 'EffectiveStartDate' to datetime
resources['EffectiveStartDate'] = pd.to_datetime(resources['Service Territory Member[EffectiveStartDate]'], errors='coerce')

# Step 2: Sort resources by 'ShopResourceKey' and 'EffectiveStartDate' (latest first)
resources_sorted = resources.sort_values(by=['ShopResourceKey', 'EffectiveStartDate'], ascending=[True, False])

# Step 3: Drop duplicates in 'resources', keeping the latest 'EffectiveStartDate' for each 'ShopResourceKey'
resources_unique = resources_sorted.drop_duplicates(subset=['ShopResourceKey'], keep='first')

# Step 4: Add 'Active' status based on date range
resources_unique['Active'] = resources_unique.apply(is_active, axis=1, args=(start_date, end_date))

# Step 5: Merge 'shifts_filtered' with 'resources_unique' (add both 'Service Resource[IsActive]' and 'Active')
shifts_filtered = shifts_filtered.merge(
    resources_unique[['ShopResourceKey', 'Service Resource[IsActive]', 'Active']],
    on='ShopResourceKey',
    how='left'
)

# Step 7: Filter for only active resources
shifts_filtered = shifts_filtered[(shifts_filtered['Service Resource[IsActive]'] == 'True') & (shifts_filtered['Active'] == True)]

# Step 9: Check the filtered data for 'PersonalNumberKey' and 'ShiftDate'
check = shifts_filtered[(shifts_filtered['PersonalNumberKey'] == '969_25367') & (pd.to_datetime(shifts_filtered['ShiftDate'], errors='coerce') == '2024-09-02')]

print(check)

shifts_filtered.columns

shifts_filtered['ShiftDurationHours'] = shifts_filtered['ShiftDurationHours'].fillna(0)
absences['Start'] = pd.to_datetime(absences['Start'], errors='coerce')
absences['End'] = pd.to_datetime(absences['End'], errors='coerce')
absences['AbsenceDurationHours'] = (absences['End'] - absences['Start']).dt.total_seconds() / 3600
absences['PersonalNumberKey'] = absences['Resource.RelatedRecord.GT_StoreCode__c'] + '_' + absences['Resource.GT_PersonalNumber__c']
absences['PersonalNumberKey'].head
# Group absences by PersonalNumberKey and date to find total absence hours per day per resource
absences.head()
absences['Start'] = absences['Start'].dt.tz_convert(None)
absences['End'] = absences['End'].dt.tz_convert(None)
absences['AbsenceDate'] = absences['Start'].dt.date

absences.head()
# Modify the filtering logic to account for absences that overlap with the start_date and end_date
absences_filtered = absences[(absences['End'] >= start_date) & (absences['Start'] <= end_date)]

def expand_multiday_absences(row):
    start_date = row['Start'].normalize()
    end_date = row['End'].normalize()
    expanded_records = []
    current_date = start_date

    # Set the time threshold (20:00)
    evening_cutoff = pd.Timestamp(current_date).replace(hour=20, minute=0, second=0)

    while current_date <= end_date:
        if current_date == start_date and current_date == end_date:
            # Absence starts and ends on the same day
            if row['Start'] > evening_cutoff:
                hours = 0  # No absence counted if start is after 20:00
            else:
                hours = (row['End'] - row['Start']).total_seconds() / 3600
        elif current_date == start_date:
            # First day: Calculate hours from start time to midnight
            if row['Start'] > evening_cutoff:
                hours = 0  # No absence counted if start is after 20:00
            else:
                end_of_day = pd.Timestamp.combine(current_date + timedelta(days=1), pd.Timestamp.min.time())  # Midnight of the next day
                hours = (end_of_day - row['Start']).total_seconds() / 3600
        elif current_date == end_date:
            # Last day: Calculate hours from midnight to the end time
            start_of_day = pd.Timestamp.combine(current_date, pd.Timestamp.min.time())  # Midnight of the current day
            hours = (row['End'] - start_of_day).total_seconds() / 3600
        else:
            # Full day: Assign 24 hours for full days within the absence period, capped at 8 hours
            hours = 24

        # Cap hours to a maximum of 8 per day
        hours = min(8, hours)
        
        expanded_records.append({
            'PersonalNumberKey': row['PersonalNumberKey'],
            'AbsenceDate': current_date.date(),
            'AbsenceDurationHours': hours,
            'AbsenceStartTime': row['Start'],
            'AbsenceEndTime': row['End'],
            'Resource.GT_PersonalNumber__c': row['Resource.GT_PersonalNumber__c'],
            'Resource.RelatedRecord.GT_StoreCode__c': row['Resource.RelatedRecord.GT_StoreCode__c'],
            'Service Resource[Id]': row['Service Resource[Id]']
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
    'Service Resource[Id]':'first',
    'AbsenceStartTime': 'first',
    'AbsenceEndTime': 'last'
}).reset_index()
absences_grouped[absences_grouped['Resource.RelatedRecord.GT_StoreCode__c'] == '005']
# Group shifts by PersonalNumberKey and ShiftDate to find total shift hours per day per resource
shifts_grouped = shifts_filtered.groupby(['PersonalNumberKey', 'ShiftDate']).agg({
    'ShiftDurationHours': 'sum',  # Sum of absence duration hours
    'GT_ServiceResource__r.Name' : 'first',
    'GT_ShopCode__c': 'first',
    'ShopResourceKey': 'first',  
    'StartDateHour': 'first',  
    'iso_year': 'first',  
    'iso_week': 'first',
    'date': 'first',
    'GT_ShopCode__c': 'first',
    'Shop[Name]': 'first',
    'StartTime': 'first',
    'EndTime': 'last'
    
}).reset_index()

check= shifts_filtered[shifts_filtered['PersonalNumberKey'] == '969_25367']
columns_to_display = ['GT_ServiceResource__r.Name', 'StartTime', 'ShiftDate', 'iso_week', 'iso_year', 
                      'StartDateHour',  'ShiftDurationHours', 'Service Resource[IsActive]', 'Active']

check_filtered = check[columns_to_display]

# Display the filtered DataFrame
check_filtered
# Merge expanded_absences with shifts data to calculate adjusted shift hours
sfshifts_merged = pd.merge(
    shifts_grouped,
    absences_grouped,
    how='left',
    left_on=['PersonalNumberKey', 'ShiftDate'],
    right_on=['PersonalNumberKey','AbsenceDate'],
    suffixes=('', '_absence')
)
# Adjust AbsenceDurationHours if ShiftDurationHours is greater
sfshifts_merged['AbsenceDurationHours'] = sfshifts_merged.apply(
    lambda row: row['ShiftDurationHours'] if row['ShiftDurationHours'] < row['AbsenceDurationHours'] else row['AbsenceDurationHours'],
    axis=1
)

# Calculate the adjusted shift duration by subtracting the absence hours
sfshifts_merged['ShiftDurationHoursAdjusted'] = sfshifts_merged['ShiftDurationHours'] - sfshifts_merged['AbsenceDurationHours'].fillna(0)
# replacing negatives to 0
sfshifts_merged['ShiftDurationHoursAdjusted'] = sfshifts_merged['ShiftDurationHoursAdjusted'].apply(lambda x: max(x, 0))
# Recalculate ShiftDurationMinutes based on adjusted hours
sfshifts_merged['ShiftDurationMinutesAdjusted'] = sfshifts_merged['ShiftDurationHoursAdjusted'] * 60
sfshifts_merged['ShiftDate'] = pd.to_datetime(sfshifts_merged['ShiftDate'])

check = sfshifts_merged[(sfshifts_merged['PersonalNumberKey'] == '017_11108') & (pd.to_datetime(sfshifts_merged['ShiftDate'], errors='coerce') == '2024-09-02')]
columns_to_display = ['ShiftDurationHours', 'AbsenceDurationHours', 'ShiftDurationHoursAdjusted']

# Filter the DataFrame and select only the required columns
check_filtered = check[columns_to_display]
check_filtered

del sfshifts, resources, absences
shift_slots = sfshifts_merged.groupby(['GT_ShopCode__c', 'Shop[Name]', 'date'])[['ShiftDurationMinutesAdjusted', 'ShiftDurationHours','AbsenceDurationHours', 'ShiftDurationHoursAdjusted']].sum().reset_index()
shift_slots['date'] = pd.to_datetime(shift_slots['date'], format='%d/%m/%Y', errors='coerce')
shift_slots['TotalSlots'] = shift_slots['ShiftDurationMinutesAdjusted'] / 5
shift_slots['TotalHours'] = shift_slots['ShiftDurationHours'].fillna(0)
shift_slots['BlockedHours'] = shift_slots['AbsenceDurationHours'].fillna(0)
shift_slots['ShiftDurationHoursAdjusted']= shift_slots['TotalHours'] - shift_slots['BlockedHours'] 

shift_slots['BlockedHoursPercentage'] = (shift_slots['BlockedHours'] / shift_slots['TotalHours']) * 100
shift_slots.head(10)
# Step 1: Generate all dates within the specified range
date_range = pd.date_range(start=start_date, end=end_date, freq='B')  # weekdays only

# Step 2: Create a DataFrame for all combinations of shop codes and the date range
shops_dates = pd.MultiIndex.from_product(
    [region_mapping['CODE'].unique(), date_range],
    names=['GT_ShopCode__c', 'date']
).to_frame(index=False)

# Step 3: Include Region, Area, and Shop[Name] information in the shops_dates DataFrame
shops_dates = pd.merge(
    shops_dates,
    region_mapping[['CODE', 'REGION', 'AREA', 'DESCR']],  # Include Region, Area, and Shop[Name]
    left_on='GT_ShopCode__c',
    right_on='CODE',
    how='left'
)

# Rename columns to match the expected output
shops_dates.rename(columns={
    'REGION': 'REGION',
    'AREA': 'AREA',
    'DESCR': 'Shop[Name]'
}, inplace=True)

# Step 4: Drop unnecessary columns
shops_dates.drop(columns=['CODE'], inplace=True)

shift_slots = pd.merge(
    shops_dates,
    shift_slots,
    on=['GT_ShopCode__c', 'date'],
    how='left', 
    suffixes=('', '_drop')  # Use '_drop' as the suffix for the columns you want to drop
)
shift_slots = shift_slots.loc[:, ~shift_slots.columns.str.endswith('_drop')]
# Step 4: Fill missing values for any shops that had no shifts
shift_slots['ShiftDurationHours'] = shift_slots['ShiftDurationHours'].fillna(0)
shift_slots['ShiftDurationMinutesAdjusted'] = shift_slots['ShiftDurationMinutesAdjusted'].fillna(0)
shift_slots['ShiftDurationHoursAdjusted'] = shift_slots['ShiftDurationHoursAdjusted'].fillna(0)
shift_slots['AbsenceDurationHours'] = shift_slots['AbsenceDurationHours'].fillna(0)
shift_slots['BlockedHours'] = shift_slots['AbsenceDurationHours'].fillna(0)-shift_slots['OverlapHours'].fillna(0)
shift_slots['BlockedHoursPercentage'] = shift_slots['BlockedHoursPercentage'].fillna(0)
shift_slots['TotalSlots'] = shift_slots['TotalSlots'].fillna(0)
shift_slots['TotalHours'] = shift_slots['TotalHours'].fillna(0)

# Test shop 'C07' to check final output
a = shift_slots[shift_slots['GT_ShopCode__c'] == '017']

# Add weekday name and ISO week
shift_slots['date'] = pd.to_datetime(shift_slots['date'], format='%d/%m/%Y')
shift_slots['day'] = shift_slots['date'].dt.day
shift_slots['weekday'] = shift_slots['date'].dt.day_name()
shift_slots['iso_week'] = shift_slots['date'].dt.isocalendar().week
shift_slots['month'] = shift_slots['date'].dt.strftime('%B')

# Remove Sundays if needed
shift_slots = shift_slots[shift_slots['weekday'] != 'Sunday']
shift_slots.columns
# Sort the shift_slots DataFrame by 'date'
shift_slots = shift_slots.sort_values(by='date')
shift_slots.rename(columns={
    'REGION': 'Region',
    'AREA': 'Area',
    'DESCR': 'Shop[Name]'
}, inplace=True)
# Save to Excel
current_date = datetime.now().strftime("%Y-%m-%d")
output_file_path = 'shiftslots_sep1.xlsx'  # Use f-string to include the date in the filename
shift_slots.to_excel(output_file_path, index=False, engine='openpyxl')

