import os
import pandas as pd

csv_path = os.environ.get('CSV_PATH', '../prod-log.csv')
df = pd.read_csv(csv_path)

# Variables - change for different month/week
month = '05_May'
week = 'Week 1'

df_filtered = df[(df['Month'] == month) & (df['Week No'] == week)]

df_filtered['Planned Hours'] = df_filtered['Planned Hours'].fillna(0)
df_filtered['Actual Hours Consumed'] = df_filtered['Actual Hours Consumed'].fillna(0)
df_filtered['Remarks'] = df_filtered['Remarks'].fillna('')

# Select columns for report
report_cols = ['Category', 'Project', 'Activity/Task', 'Planned Hours', 'Actual Hours Consumed', 'Remarks']
df_report = df_filtered[report_cols].copy()

# Group by Category and print report
for category in df_report['Category'].unique():
    cat_data = df_report[df_report['Category'] == category]
    
    planned_hours = cat_data['Planned Hours']
    actual_hours = cat_data['Actual Hours Consumed']
    variance = planned_hours - actual_hours

    print(f"Category: {category}")
    print(f"{'='*80}")
    print(cat_data[['Project', 'Planned Hours', 'Actual Hours Consumed', 'Remarks']].to_string(index=False))
    print(f"Total: Planned={planned_hours.sum()}, Actual={actual_hours.sum()}, Variance={variance.sum()}")
