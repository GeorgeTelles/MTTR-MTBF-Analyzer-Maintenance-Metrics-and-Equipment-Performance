import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# Load data from the Excel file generated in the previous code
equipments_df = pd.read_excel('Equipment_and_Failures_Refinery.xlsx', sheet_name='Equipments')
failures_df = pd.read_excel('Equipment_and_Failures_Refinery.xlsx', sheet_name='Equipment_Failures')
preventive_maintenance_df = pd.read_excel('Equipment_and_Failures_Refinery.xlsx', sheet_name='Preventive_Maintenance')

# Convert date columns to datetime format
failures_df['Failure_Date'] = pd.to_datetime(failures_df['Failure_Date'])
failures_df['Resolution_Date'] = pd.to_datetime(failures_df['Resolution_Date'])
failures_df = failures_df.sort_values(by='Failure_Date', ascending=True)

# Calculate downtime and repair time
failures_df['Downtime_Hours'] = (failures_df['Resolution_Date'] - failures_df['Failure_Date']).dt.total_seconds() / 3600
failures_df['Repair_Time_Hours'] = failures_df['Downtime_Hours']  # Assuming repair time equals downtime

# Calculate MTTR (Mean Time To Repair)
mttr_per_equipment = failures_df.groupby('Equipment_ID')['Repair_Time_Hours'].mean().reset_index()
mttr_per_equipment.columns = ['Equipment_ID', 'MTTR_Hours']

# Calculate MTBF (Mean Time Between Failures)
failures_df['Time_Between_Failures_Hours'] = failures_df.groupby('Equipment_ID')['Failure_Date'].diff().dt.total_seconds() / 3600
mtbf_per_equipment = failures_df.groupby('Equipment_ID')['Time_Between_Failures_Hours'].mean().reset_index()
mtbf_per_equipment.columns = ['Equipment_ID', 'MTBF_Hours']

# Merge MTTR and MTBF metrics
metric_analysis = pd.merge(mttr_per_equipment, mtbf_per_equipment, on='Equipment_ID')

# Add equipment details (name and type) to the analysis
metric_analysis = pd.merge(metric_analysis, equipments_df[['Equipment_ID', 'Equipment_Name', 'Type']], on='Equipment_ID')

print(metric_analysis)

# MTTR Bar Chart
fig_mttr = px.bar(metric_analysis, x='MTTR_Hours', y='Equipment_Name',
                  title='Mean Time To Repair (MTTR) by Equipment',
                  labels={'MTTR_Hours': 'Repair Time (Hours)', 'Equipment_Name': 'Equipment'},
                  color_discrete_sequence=px.colors.qualitative.Set1)
fig_mttr.update_layout(barmode='group', height=800)

# MTBF Bar Chart
fig_mtbf = px.bar(metric_analysis, x='MTBF_Hours', y='Equipment_Name', 
                  title='Mean Time Between Failures (MTBF) by Equipment',
                  labels={'MTBF_Hours': 'Time Between Failures (Hours)', 'Equipment_Name': 'Equipment'},
                  color_discrete_sequence=px.colors.sequential.Viridis)
fig_mtbf.update_layout(barmode='group', height=800)

fig_mttr.show()
fig_mtbf.show()

# Maintenance Cost Analysis
print(failures_df['Maintenance_Cost'].head())

# Clean the Maintenance_Cost column by removing '$', commas, and '00'
failures_df['Maintenance_Cost'] = failures_df['Maintenance_Cost'].apply(
    lambda x: str(x).replace('$', '').replace(',', '').replace('00', '').strip())
# Convert the column to float after cleaning
failures_df['Maintenance_Cost'] = failures_df['Maintenance_Cost'].astype(float)

# Verify the column after conversion
print(failures_df['Maintenance_Cost'].head())

# Calculate total maintenance cost per equipment
maintenance_cost_per_equipment = failures_df.groupby('Equipment_ID')['Maintenance_Cost'].sum().reset_index()
maintenance_cost_per_equipment = pd.merge(maintenance_cost_per_equipment, equipments_df[['Equipment_ID', 'Equipment_Name']], on='Equipment_ID')

# Plot maintenance cost bar chart
fig_maintenance_cost = px.bar(maintenance_cost_per_equipment, x='Maintenance_Cost', y='Equipment_Name', 
                              title='Total Maintenance Cost by Equipment',
                              labels={'Maintenance_Cost': 'Maintenance Cost ($)', 'Equipment_Name': 'Equipment'},
                              color_discrete_sequence=px.colors.sequential.Viridis)
fig_maintenance_cost.update_layout(height=800)

fig_maintenance_cost.show()

# Failure Cause Analysis
most_frequent_failure_causes = failures_df['Failure_Cause'].value_counts().reset_index()
most_frequent_failure_causes.columns = ['Failure_Cause', 'Count']

# Plot failure cause bar chart
fig_failure_causes = px.bar(most_frequent_failure_causes, x='Count', y='Failure_Cause', 
                            title='Most Frequent Failure Causes',
                            labels={'Count': 'Number of Occurrences', 'Failure_Cause': 'Failure Cause'},
                            color_discrete_sequence=px.colors.qualitative.Set2)
fig_failure_causes.update_layout(height=800)

fig_failure_causes.show()

# Top 5 equipment with highest MTBF
top_5_mtbf = metric_analysis.nlargest(5, 'MTBF_Hours')

# Top 5 equipment with highest MTTR
top_5_mttr = metric_analysis.nlargest(5, 'MTTR_Hours')

# Top 5 equipment with lowest MTBF
bottom_5_mtbf = metric_analysis.nsmallest(5, 'MTBF_Hours')

# Top 5 equipment with lowest MTTR
bottom_5_mttr = metric_analysis.nsmallest(5, 'MTTR_Hours')

# Plot charts

# MTBF chart for top 5
fig_top_5_mtbf = px.bar(top_5_mtbf, x='MTBF_Hours', y='Equipment_Name', color='Equipment_Name', 
                        title='Top 5 Equipment with Highest MTBF - Mean Time Between Failures',
                        labels={'MTBF_Hours': 'Time Between Failures (Hours)', 'Equipment_Name': 'Equipment'},
                        color_discrete_sequence=px.colors.sequential.Plasma)
fig_top_5_mtbf.update_layout(height=800)

# MTTR chart for top 5
fig_top_5_mttr = px.bar(top_5_mttr, x='MTTR_Hours', y='Equipment_Name', color='Equipment_Name', 
                         title='Top 5 Equipment with Highest MTTR - Mean Time To Repair',
                         labels={'MTTR_Hours': 'Repair Time (Hours)', 'Equipment_Name': 'Equipment'},
                         color_discrete_sequence=px.colors.sequential.Viridis)
fig_top_5_mttr.update_layout(height=800)

# MTBF chart for bottom 5
fig_bottom_5_mtbf = px.bar(bottom_5_mtbf, x='MTBF_Hours', y='Equipment_Name', color='Equipment_Name', 
                           title='Top 5 Equipment with Lowest MTBF - Mean Time Between Failures',
                           labels={'MTBF_Hours': 'Time Between Failures (Hours)', 'Equipment_Name': 'Equipment'},
                           color_discrete_sequence=px.colors.sequential.RdBu)
fig_bottom_5_mtbf.update_layout(height=800)

# MTTR chart for bottom 5
fig_bottom_5_mttr = px.bar(bottom_5_mttr, x='MTTR_Hours', y='Equipment_Name', color='Equipment_Name', 
                           title='Top 5 Equipment with Lowest MTTR - Mean Time To Repair',
                           labels={'MTTR_Hours': 'Repair Time (Hours)', 'Equipment_Name': 'Equipment'},
                           color_discrete_sequence=px.colors.sequential.Viridis)
fig_bottom_5_mttr.update_layout(height=800)

fig_top_5_mtbf.show()
fig_top_5_mttr.show()
fig_bottom_5_mtbf.show()
fig_bottom_5_mttr.show()