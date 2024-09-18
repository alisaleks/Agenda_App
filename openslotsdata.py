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
'Resource Absence[AbsenceNumber]': str,
'Service Resource[Name]': str,
'Service Resource[Id]':str
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
        'Service Territory Member[ServiceResourceId]', 'Service Resource[GT_PersonalNumber__c]', 'Service Resource[IsActive]',
        'Service Resource[GT_Role__c]'
    ], 
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
        'Service Resource[GT_PersonalNumber__c]', 'User[GT_StoreCode__c]', 'Service Resource[Id]'
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
region_mapping.columns
# Filter the original region_mapping DataFrame
region_mapping = region_mapping[region_mapping['SYM'] != 'N']

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
print(shifts_filtered[shifts_filtered['GT_ShopCode__c'] == '978'])

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
shifts_filtered = shifts_filtered.sort_values(by=['Key', 'LastModifiedDate'], ascending=[True, False])
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

# Step 6: Create a new key combining 'ShopResourceKey' and 'Shift[ShiftNumber]' to differentiate shifts
shifts_filtered['UniqueShiftKey'] = shifts_filtered['ShopResourceKey'] + '_' + shifts_filtered['Shift[ShiftNumber]']

# Step 7: Filter for only active resources
shifts_filtered = shifts_filtered[(shifts_filtered['Service Resource[IsActive]'] == 'True') & (shifts_filtered['Active'] == True)]

# Step 8: Remove duplicates based on the 'UniqueShiftKey'
shifts_filtered = shifts_filtered.drop_duplicates(subset=['UniqueShiftKey'])

# Step 9: Check the filtered data for 'PersonalNumberKey' and 'ShiftDate'
check = shifts_filtered[(shifts_filtered['PersonalNumberKey'] == '969_25367') & (pd.to_datetime(shifts_filtered['ShiftDate'], errors='coerce') == '2024-09-02')]

print(check)

shifts_filtered.columns

shifts_filtered['ShiftDurationHours'] = shifts_filtered['ShiftDurationHours'].fillna(0)

absences['Start'] = pd.to_datetime(absences['Start'], errors='coerce')
absences['End'] = pd.to_datetime(absences['End'], errors='coerce')
absences['AbsenceDurationHours'] = (absences['End'] - absences['Start']).dt.total_seconds() / 3600
absences['PersonalNumberKey'] = absences['Resource.RelatedRecord.GT_StoreCode__c'] + '_' + absences['Resource.GT_PersonalNumber__c']

# Group absences by PersonalNumberKey and date to find total absence hours per day per resource
absences['AbsenceDate'] = absences['Start'].dt.date
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
            'AbsenceNumber': row['AbsenceNumber'],
            'Resource.GT_PersonalNumber__c': row['Resource.GT_PersonalNumber__c'],
            'Resource.RelatedRecord.GT_StoreCode__c': row['Resource.RelatedRecord.GT_StoreCode__c'],
            'Resource.Name': row['Resource.Name'],
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
    'AbsenceNumber': 'first',  
    'Resource.Name': 'first',  
    'AbsenceStartTime': 'first',
    'AbsenceEndTime': 'last'
}).reset_index()
absences_grouped[absences_grouped['Resource.RelatedRecord.GT_StoreCode__c'] == 'A71']
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

# Filter appointments within August
appointments_filtered = appointments[(appointments['ApptStartTime'] >= start_date) & (appointments['ApptEndTime'] <= end_date)].copy()

#Deleting unused DataFrames to free up memory
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
        slots.append((row['Service Appointment[GT_ShopCode__c]'], row['Service Resource[Name]'], slot_start.strftime('%d/%m/%Y'), slot_start, row['ApptsLastModifiedDate'], row['ApptStartTime'], 
            row['ApptEndTime'], row['Service Appointment[GT_ServiceResource__c]']))
        slot_start += timedelta(minutes=5)

# Create a DataFrame from all the slots including the last modified date for duplicate removal
all_slots_df = pd.DataFrame(slots, columns=['GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot', 'LastModifiedDate', 'ApptStartTime', 'ApptEndTime', 'Service Appointment[GT_ServiceResource__c]'])
# Ensure the 'date' column is of datetime type for proper merging
all_slots_df['date'] = pd.to_datetime(all_slots_df['date'], format='%d/%m/%Y')
all_slots_df['PersonalidKey'] = all_slots_df['GT_ShopCode__c'] + all_slots_df['Service Appointment[GT_ServiceResource__c]']
# Sort slots by 'GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot', and 'LastModifiedDate' to prioritize keeping the earliest slot
all_slots_df = all_slots_df.sort_values(by=['GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot', 'LastModifiedDate'], ascending=[True, True, True, True, False])

# Remove duplicate 5-minute slots, keeping only the earliest modified slot
all_slots_df = all_slots_df.drop_duplicates(subset=['GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot'], keep='first')

# Step 1: Initialize an empty list to store expanded shifts
shift_slots_5mins = []

# Step 2: Loop through each row in sfshifts_merged and generate 5-minute slots
for _, row in sfshifts_merged.iterrows():
    shift_start = row['StartTime']
    shift_end = row['EndTime']
    
    # Generate 5-minute slots between shift_start and shift_end
    while shift_start < shift_end:
        shift_slots_5mins.append({
            'ShopResourceKey': row['ShopResourceKey'],
            'PersonalNumberKey': row['PersonalNumberKey'],
            'ShiftSlot': shift_start,  # The start time of the slot
            'ShiftDurationHoursAdjusted': row['ShiftDurationHoursAdjusted'],
            'ShiftDurationMinutesAdjusted': row['ShiftDurationMinutesAdjusted'],
            'AbsenceDurationHours': row['AbsenceDurationHours'],
            'ShiftLabel': row['Shift[Label]'],
            'GT_ServiceResource__r.Name': row['GT_ServiceResource__r.Name'], 
            'StartTime': row['StartTime'], 
            'EndTime': row['EndTime']
        })
        shift_start += timedelta(minutes=5)

# Step 3: Convert the list to a DataFrame
expanded_shifts_df = pd.DataFrame(shift_slots_5mins)

# Assuming 'StartTime' and 'EndTime' columns are added during the slot expansion process.
filtered_shifts_df = expanded_shifts_df[['ShiftSlot', 'StartTime', 'EndTime', 'ShiftLabel']]

absence_slots_5mins = []

# Step 2: Loop through each row in absences_grouped and generate 5-minute slots
for _, row in absences_grouped.iterrows():
    absence_start = pd.Timestamp.combine(row['AbsenceDate'], row['AbsenceStartTime'].time()) if pd.notnull(row['AbsenceStartTime']) else row['AbsenceDate']
    absence_end = pd.Timestamp.combine(row['AbsenceDate'], row['AbsenceEndTime'].time()) if pd.notnull(row['AbsenceEndTime']) else row['AbsenceDate'] + timedelta(hours=8)
    
    # Generate 5-minute slots between absence_start and absence_end
    while absence_start < absence_end:
        absence_slots_5mins.append({
            'PersonalNumberKey': row['PersonalNumberKey'],
            'AbsenceSlot': absence_start,  # The start time of the absence slot
            'AbsenceDurationHours': row['AbsenceDurationHours'],
            'AbsenceNumber': row['AbsenceNumber'],
            'Resource.GT_PersonalNumber__c': row['Resource.GT_PersonalNumber__c'], 
            'GT_ShopCode__c':row['Resource.RelatedRecord.GT_StoreCode__c'],
            'Resource.Name': row['Resource.Name'],
            'Service Resource[Id]': row['Service Resource[Id]']
        })
        absence_start += timedelta(minutes=5)

# Step 3: Convert the list to a DataFrame
expanded_absences_df = pd.DataFrame(absence_slots_5mins)
expanded_absences_df['PersonalidKey'] = expanded_absences_df['GT_ShopCode__c'] + expanded_absences_df['Service Resource[Id]']

# Step 1: Merge shift slots and appointment slots based on PersonalidKey, date, and Slot
# The goal here is to ensure only appointments that match an active shift are kept.
appointments_with_shifts = pd.merge(
    all_slots_df,
    expanded_shifts_df[['ShopResourceKey', 'ShiftSlot']],
    how='inner',
    left_on=['PersonalidKey', 'Slot'],
    right_on=['ShopResourceKey', 'ShiftSlot']
)

appointments_with_shifts

target_date = pd.to_datetime('2024-09-02')

# For simplicity, let's assume '001_0Hn6700000001OXCAY' is the PersonalidKey for both shifts and absences
personal_key = '94B0Hn6700000001PBCAY'

# Filter appointments, shifts, and absences for '001_0Hn6700000001OXCAY' and '2024-09-06'
appointments_example = all_slots_df[
    (all_slots_df['PersonalidKey'] == personal_key) & 
    (all_slots_df['date'] == target_date)
]



absences_example = expanded_absences_df[
    (expanded_absences_df['PersonalidKey'] == personal_key) & 
    (expanded_absences_df['AbsenceSlot'].dt.date == target_date.date())
]

# Display a few rows of the filtered datasets to understand the data for this specific PersonalidKey and date
print(appointments_example.head(20))
print(absences_example.head(20))

# Let's also ensure that absence does not block the appointment slot.

# Step 3: Merge with absence slots to check for overlaps
appointments_with_shifts_and_absences = pd.merge(
    appointments_with_shifts,
    expanded_absences_df[['PersonalidKey', 'AbsenceSlot']],
    how='left',
    left_on=['PersonalidKey', 'Slot'],
    right_on=['PersonalidKey', 'AbsenceSlot']
)
# Step 4: Prioritize appointments over absences
# If there is an overlap between an appointment and an absence, we prioritize the appointment by removing the absence overlap.
appointments_with_shifts_and_absences.columns

appointments_with_shifts_and_absences['IsAbsence'] = appointments_with_shifts_and_absences['AbsenceSlot'].notnull()
appointments_with_shifts_and_absences['IsAppointment'] = appointments_with_shifts_and_absences['Slot'].notnull()
# expanded_absences_df already contains all absences, so we'll use it directly
all_absence_slots = expanded_absences_df.copy()

# Merge absence slots with appointment slots to identify overlaps
overlapping_absence_slots = pd.merge(
    expanded_absences_df[['PersonalidKey', 'AbsenceSlot']],
    appointments_with_shifts_and_absences[['PersonalidKey', 'Slot']],
    left_on=['PersonalidKey', 'AbsenceSlot'],
    right_on=['PersonalidKey', 'Slot'],
    how='inner'
)

# Only keep the necessary columns
overlapping_absence_slots = overlapping_absence_slots[['PersonalidKey', 'AbsenceSlot']]
# Add GT_ShopCode__c and AbsenceSlot date to the overlapping absence slots for grouping
overlapping_absence_slots = pd.merge(
    overlapping_absence_slots,
    expanded_absences_df[['PersonalidKey', 'GT_ShopCode__c', 'AbsenceSlot']],
    on=['PersonalidKey', 'AbsenceSlot'],
    how='left'
)

# Convert AbsenceSlot to date for grouping purposes
overlapping_absence_slots['AbsenceSlotDate'] = overlapping_absence_slots['AbsenceSlot'].dt.date

# Group by shop, service resource, and date to calculate the total overlapping absence slots
total_overlapping_absence_slots = overlapping_absence_slots.groupby(
    ['GT_ShopCode__c', 'AbsenceSlotDate']
).size().reset_index(name='TotalOverlappingAbsenceSlots')

# Rename the AbsenceSlotDate to 'date' to align with other datasets
total_overlapping_absence_slots.rename(columns={'AbsenceSlotDate': 'date'}, inplace=True)

# Step 2: Count the unique slots, ensuring no double-counting for overlaps
net_booked_slots = appointments_with_shifts_and_absences.groupby(['GT_ShopCode__c', 'Service Resource[Name]', 'date', 'Slot']).size().reset_index(name='Count')

# Normalize counts to 1 for any slot count greater than 1
net_booked_slots['Count'] = net_booked_slots['Count'].apply(lambda x: 1 if x > 0 else 0)

# Calculate the total booked slots by shop, service resource, and date
total_booked_slots_by_date = net_booked_slots.groupby(['GT_ShopCode__c', 'Service Resource[Name]', 'date'])['Count'].sum().reset_index()
total_booked_slots_by_date.head()
total_booked_slots_by_date['date'] = pd.to_datetime(total_booked_slots_by_date['date'], errors='coerce')
grouped_df = total_booked_slots_by_date.groupby(['GT_ShopCode__c', 'date']).agg({'Count': 'sum'}).reset_index()
# Calculate the total number of 5-minute slots available per shop and date
total_overlapping_absence_slots['date'] = pd.to_datetime(total_overlapping_absence_slots['date'], errors='coerce')
shift_slots = sfshifts_merged.groupby(['GT_ShopCode__c', 'Shop[Name]', 'date'])[['ShiftDurationMinutesAdjusted', 'ShiftDurationHours','AbsenceDurationHours', 'ShiftDurationHoursAdjusted']].sum().reset_index()
shift_slots['date'] = pd.to_datetime(shift_slots['date'], format='%d/%m/%Y', errors='coerce')
shift_slots = pd.merge(shift_slots, total_overlapping_absence_slots, on=['GT_ShopCode__c', 'date'], how='left')
shift_slots['TotalOverlappingAbsenceSlots'] = shift_slots['TotalOverlappingAbsenceSlots'].fillna(0)
shift_slots['OverlapHours'] = (shift_slots['TotalOverlappingAbsenceSlots']* 5) / 60
shift_slots['TotalSlots'] = shift_slots['ShiftDurationMinutesAdjusted'] / 5
shift_slots['TotalHours'] = shift_slots['ShiftDurationHours'].fillna(0)
shift_slots['BlockedHours'] = shift_slots['AbsenceDurationHours'].fillna(0)-shift_slots['OverlapHours'].fillna(0)
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


# Step 5: Recalculate `TotalBookedSlots` based on the total booked slots by date
shift_slots = pd.merge(
    shift_slots, 
    grouped_df, 
    on=['GT_ShopCode__c', 'date'], 
    how='left'
)

# Fill missing values for TotalBookedSlots and perform necessary calculations
shift_slots['TotalBookedSlots'] = shift_slots['Count'].fillna(0)
shift_slots.drop(columns=['Count'], inplace=True)

shift_slots['OpenSlots'] = shift_slots['TotalSlots'] - shift_slots['TotalBookedSlots']
shift_slots['OpenSlots'] = shift_slots['OpenSlots'].apply(lambda x: max(x, 0))

# Step 6: Additional Calculations
shift_slots['OpenHours'] = (shift_slots['OpenSlots'] * 5) / 60
shift_slots['BookedHours'] = (shift_slots['TotalBookedSlots'] * 5) / 60
shift_slots['SaturationPercentage'] = (shift_slots['BookedHours'] / shift_slots['TotalHours']) * 100
shift_slots['SaturationPercentage'] = shift_slots['SaturationPercentage'].clip(lower=0, upper=100)

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
output_file_path = f'shiftslots_{current_date}.xlsx'  # Use f-string to include the date in the filename
shift_slots.to_excel(output_file_path, index=False, engine='openpyxl')



#TAB4
sfshifts_merged.head()
# Step 3: Include Region, Area, and Shop[Name] information in the shops_dates DataFrame
sfshifts_merged = pd.merge(
    sfshifts_merged,
    region_mapping[['CODE', 'REGION', 'AREA', 'DESCR']],  
    left_on='GT_ShopCode__c',
    right_on='CODE',
    how='left'
)

missing_shop_codes = sfshifts_merged[sfshifts_merged['GT_ShopCode__c'].isna()]
print(missing_shop_codes[['GT_ShopCode__c', 'Shop[Name]', 'REGION', 'AREA', 'CODE', 'ShiftDurationHoursAdjusted']])

# Rename columns to match the expected output
sfshifts_merged.rename(columns={
    'REGION': 'Region',
    'AREA': 'Area',
    'DESCR': 'Shop[Name]'
}, inplace=True)

# Step 4: Drop unnecessary columns
sfshifts_merged.drop(columns=['CODE'], inplace=True)
sfshifts_merged.fillna(0, inplace=True)

sfshifts_merged['weekday'] = sfshifts_merged['ShiftDate'].dt.day_name()
sfshifts_merged.columns

# Save to Excel
output_file_path2 = 'hcpshiftslots.xlsx'
sfshifts_merged.to_excel(output_file_path2, index=False, engine='openpyxl')


# Step 1: Read and prepare HCMdata
hcm_file = 'HCMShifts.csv'
hcm_columns_to_string = {
    'Shop[Shop Code - Descr]': str,
    'Unique Employee[Employee Full Name]': str,
    'Unique Employee[Employee Person Number]': str
}

HCMdata = pd.read_csv(hcm_file, engine='python', dtype=hcm_columns_to_string, usecols=[
    'Shop[Shop Code - Descr]', 'Unique Employee[Employee Full Name]', 'Unique Employee[Employee Person Number]',
    'Calendar[ISO Week]', 'Calendar[ISO Year]', '[Audiologist_FTE]'
])

def get_previous_weeks_range(n=2, future_weeks=3):
    today = datetime.today()
    current_iso_year, current_iso_week, _ = today.isocalendar()
    
    # Calculate the start ISO week based on the range (n weeks back)
    start_iso_week = max(1, current_iso_week - n)
    
    # Calculate the future date (end week) by adding `future_weeks` to the current date
    future_date = today + timedelta(weeks=future_weeks)
    end_iso_year, end_iso_week, _ = future_date.isocalendar()
    
    return start_iso_week, end_iso_week, current_iso_year, end_iso_year

# Example usage
start_iso_week, end_iso_week, current_iso_year, end_iso_year = get_previous_weeks_range()
print(f"Start ISO week: {start_iso_week}, End ISO week: {end_iso_week}, Current ISO year: {current_iso_year}, End ISO year: {end_iso_year}")

# Filter HCMdata between start_iso_week and end_iso_week without considering the year
HCMdata = HCMdata[
    (HCMdata['Calendar[ISO Week]'] >= start_iso_week) &
    (HCMdata['Calendar[ISO Week]'] <= end_iso_week)
]


# Create composite keys
HCMdata['ShopCode_3char'] = HCMdata['Shop[Shop Code - Descr]'].str[:3]
HCMdata['CompositeKey'] = HCMdata['ShopCode_3char'] + '_' + HCMdata['Unique Employee[Employee Person Number]'].astype(str) + '_' + HCMdata['Calendar[ISO Year]'].astype(str) + '_' + HCMdata['Calendar[ISO Week]'].astype(str)
sfshifts_merged['CompositeKey'] = sfshifts_merged['PersonalNumberKey'] + '_' + sfshifts_merged['iso_year'].astype(str) + '_' + sfshifts_merged['iso_week'].astype(str)

# Step 2: Group and sum data
HCMdata_summed = HCMdata.groupby(
    ['CompositeKey', 'Calendar[ISO Year]', 'Calendar[ISO Week]']
).agg({
    '[Audiologist_FTE]': 'sum'
}).reset_index()

# Multiply the '[Audiologist_FTE]' by 40 to get the duration
HCMdata_summed['Duración HCM'] = HCMdata_summed['[Audiologist_FTE]'] * 40
# Step 3: Process SF shifts data
shift_duration_per_week = sfshifts_merged.groupby(
    ['CompositeKey']
).agg({
    'ShiftDurationHours': 'sum'
}).reset_index()

shift_duration_per_week.rename(columns={'ShiftDurationHours': 'Duración SF'}, inplace=True)

# Step 4: Merge both datasets (without region/area/shop data yet)
all_composite_keys = pd.merge(
    shift_duration_per_week[['CompositeKey', 'Duración SF']],  # From SF shifts
    HCMdata_summed[['CompositeKey', 'Duración HCM']],  # From HCM data
    on='CompositeKey', how='outer'
)

# Step 5: Add region, area, and shop (DESCR) mapping data based on the merged composite keys
all_composite_keys['ShopCode_3char'] = all_composite_keys['CompositeKey'].str[:3]  # Extract the ShopCode_3char from CompositeKey
all_composite_keys['iso_week'] = all_composite_keys['CompositeKey'].apply(lambda x: x.split('_')[-1])
# Merge with the region_mapping to add the 'REGION', 'AREA', and 'DESCR' (Shop Name)
all_composite_keys = pd.merge(
    all_composite_keys,
    region_mapping[['CODE', 'REGION', 'AREA', 'DESCR']],  # Add the region and area mapping
    left_on='ShopCode_3char',
    right_on='CODE',
    how='left'
)
all_composite_keys.head()
# Step 6: Final Calculations and Fill Missing Values
all_composite_keys['Diferencia de duración'] = all_composite_keys['Duración HCM'].fillna(0) - all_composite_keys['Duración SF'].fillna(0)
missing_region_rows = all_composite_keys[all_composite_keys['REGION'].isna()]
print(missing_region_rows)

# Step 7: Final output structure and rename columns
all_composite_keys = all_composite_keys[[
    'CompositeKey', 'Duración SF', 'Duración HCM', 'Diferencia de duración', 'REGION', 'AREA', 'DESCR', 'iso_week'
]]

# Rename for clarity
all_composite_keys.rename(columns={
    'CompositeKey': 'Clave compuesta',
    'REGION': 'REGION',
    'AREA': 'AREA',
    'DESCR': 'Shop Name',
}, inplace=True)
# Remove rows where REGION is blank (i.e., NaN)
all_composite_keys = all_composite_keys[all_composite_keys['REGION'].notna()]

# Fill NaN values in the following columns with 0
all_composite_keys[['Duración SF', 'Duración HCM', 'Diferencia de duración']] = all_composite_keys[['Duración SF', 'Duración HCM', 'Diferencia de duración']].fillna(0)

# Step 7: Save the result to Excel
output_file_path1 = 'hcm_sf_merged.xlsx'
all_composite_keys.to_excel(output_file_path1, index=False, engine='openpyxl')
