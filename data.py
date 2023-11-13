import pandas as pd
from datetime import datetime

df = pd.read_excel("C:\\Users\\Evence Langa\\OneDrive - Net Nine Nine\\Desktop\\Events Report\\Events Last Week_1.xls")

# Convert 'Date And Time' column to datetime data type, handling invalid values
df['Date And Time'] = pd.to_datetime(df['Date And Time'], errors='coerce')

df = df.dropna(subset=['Date And Time'])

df['Day'] = df['Date And Time'].dt.strftime('%A')

df['Full Name'] = df['First Name'] + ' ' + df['Last Name']

# Custom functions to handle first check-in and last check-out when there's no data
def first_check_in(x):
    return x.min() if not x.empty else pd.NaT

def last_check_out(x):
    return x.max() if not x.empty else pd.NaT

grouped = df.groupby(['Full Name', 'Day'])

result_df = grouped.agg({'Date And Time': [first_check_in, last_check_out]}).reset_index()

# Ensure the 'Day' column is ordered correctly
weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
result_df['Day'] = pd.Categorical(result_df['Day'], categories=weekday_order, ordered=True)

# Reset the column names as needed
result_df.columns = ['Full Name', 'Day', 'First Check-in', 'Last Check-out']

# Convert columns back to appropriate data types
result_df['First Check-in'] = pd.to_datetime(result_df['First Check-in'], errors='coerce')
result_df['Last Check-out'] = pd.to_datetime(result_df['Last Check-out'], errors='coerce')

# Fill NaN values with 0 for 'First Check-in' and 'Last Check-out'
result_df[['First Check-in', 'Last Check-out']] = result_df[['First Check-in', 'Last Check-out']].fillna(0)

# Create a DataFrame with all combinations of 'Full Name' and 'Day'
all_combinations = pd.MultiIndex.from_product([result_df['Full Name'].unique(), weekday_order], names=['Full Name', 'Day']).to_frame(index=False)
all_combinations = all_combinations.merge(result_df, how='left')

# Fill NaN values with 0 for 'First Check-in' and 'Last Check-out' in the new DataFrame
all_combinations[['First Check-in', 'Last Check-out']] = all_combinations[['First Check-in', 'Last Check-out']].fillna(0)

all_combinations.to_excel('check_ins_out_summary.xlsx', index=False)
