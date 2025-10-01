import pandas as pd
from utils import load_data, save_to_excel, OUTPUT_DIR
data = load_data("NPS_dataset.xlsx")
critical_columns = ['Ticket No','Created Date','Created Time','Resolved Date','Resolved Time','Assigned To','Status','Sub Status','Phone','Email']
missing_info = data[critical_columns].isna().sum()
status_issues = data[(data['Status']=='Complete') & (data['Sub Status']=='Pending')]
duplicate_records = data[data.duplicated(subset=['Ticket No'], keep=False)]
task1_summary = pd.DataFrame({
    'Missing_Info':[missing_info.sum()],
    'Status_Inconsistencies':[len(status_issues)],
    'Duplicate_Records':[len(duplicate_records)]
})
missing_details = data[data[critical_columns].isna().any(axis=1)].copy()
for col in critical_columns:
    missing_details[col + "_Missing"] = missing_details[col].isna()
recommendations = pd.DataFrame({
    "Recommendation":[
        "Make 'Ticket No' a mandatory unique field",
        "Ensure 'Created/Resolved Date and Time' and 'Assigned To' are always filled",
        "Enforce Status/Sub Status consistency rules",
        "Use dropdowns for 'Status' and 'Sub Status' to standardize entries",
        "Mandate Phone and Email for follow-up communication",
        "Provide a ticket submission template for all programs",
        "Implement validation at data entry to prevent missing critical fields"
    ],
    "Purpose/Impact":[
        "Prevents duplicate tickets and ensures unique identification",
        "Ensures full traceability from creation to resolution",
        "Maintains logical consistency in workflow",
        "Reduces manual errors and standardizes inputs",
        "Ensures customer contactability and better service",
        "Ensures consistent data collection across teams",
        "Reduces future data quality issues and improves reporting accuracy"
    ]
})
data['Created_DateTime'] = pd.to_datetime(data['Created Date'] + ' ' + data['Created Time'])
data['Resolved_DateTime'] = pd.to_datetime(data['Resolved Date'] + ' ' + data['Resolved Time'])
data['Resolution Time (days)'] = (data['Resolved_DateTime'] - data['Created_DateTime']).dt.days
top5_issues_freq = data['Issue 2 - NPS'].value_counts().head(5).to_frame("Frequency")
impact_scores = data.groupby('Issue 2 - NPS').agg(
    Frequency=('Ticket No','count'),
    AvgResolution=('Resolution Time (days)','mean'),
    AvgNPS=('NPS Rating','mean')
).reset_index()
impact_scores['ImpactScore'] = impact_scores['Frequency']*0.5 + impact_scores['AvgResolution']*0.3 + impact_scores['AvgNPS']*0.2
top5_issues_impact = impact_scores.sort_values('ImpactScore',ascending=False).head(5)
cross_program_issues = data.groupby(['Issue 2 - NPS','Program Name']).size().unstack(fill_value=0).reset_index()
resolution_analysis = data.groupby('Issue 2 - NPS')['Resolution Time (days)'].mean().to_frame("AvgResolutionTime").reset_index()
single_fix = top5_issues_impact.head(1)
data['Week'] = data['Created_DateTime'].dt.isocalendar().week
data['Month'] = data['Created_DateTime'].dt.month
data['Quarter'] = data['Created_DateTime'].dt.to_period('Q')
weekly_trends = data.groupby('Week').size().to_frame("WeeklyTickets").reset_index()
monthly_trends = data.groupby('Month').size().to_frame("MonthlyTickets").reset_index()
quarterly_trends = data.groupby('Quarter').size().to_frame("QuarterlyTickets").reset_index()
issue_distribution = data.groupby(['Program Name','Issue 2 - NPS']).size().unstack(fill_value=0).reset_index()
team_performance = data.groupby('Assigned To')['Resolution Time (days)'].mean().to_frame("TeamAvgResolution").reset_index()
top3_issues = data['Issue 2 - NPS'].value_counts().head(3).index
root_cause_data = data[data['Issue 2 - NPS'].isin(top3_issues)]
root_cause_summary = root_cause_data.groupby('Issue 2 - NPS')['Ticket All Remarks'].apply(lambda x: ' | '.join(x)).reset_index()
systemic_analysis = root_cause_data.groupby('Issue 2 - NPS')['Ticket No'].count().reset_index().rename(columns={'Ticket No': 'Number of Tickets'})
process_mapping = root_cause_data.groupby('Issue 2 - NPS')['Disposition Folder Level 1'].apply(lambda x: ', '.join(x.unique())).reset_index()
preventive_measures = pd.DataFrame({
    'Issue 2 - NPS': top3_issues,
    'Preventive Measures': [
        "Standardize ticket submission; enforce mandatory fields",
        "Train staff on common errors to reduce repeated issues",
        "Automate routing of tickets to correct operational process"
    ]
})
kpi_metrics = pd.DataFrame({
    "Metric": [
        'Total Tickets',
        'Avg Resolution Time (days)',
        'Pending Tickets',
        'Completed Tickets',
        'Tickets per Week',
        'Tickets per Month'
    ],
    "Value": [
        len(data),
        data['Resolution Time (days)'].mean(),
        len(data[data['Status'] == 'Pending']),
        len(data[data['Status'] == 'Complete']),
        data.groupby(data['Created_DateTime'].dt.isocalendar().week)['Ticket No'].count().mean(),
        data.groupby(data['Created_DateTime'].dt.month)['Ticket No'].count().mean()
    ]
})
resource_recommendations = pd.DataFrame({
    'Recommendation': [
        'Automate repetitive ticket validation and reporting',
        'Weekly review for operational managers, monthly for program leads',
        'Define roles: Support staff -> Team Lead -> Program Manager',
        'Implement feedback loops from resolved tickets to training team'
    ],
    'Purpose/Impact': [
        'Reduces manual errors and saves time',
        'Ensures timely oversight at multiple levels',
        'Clear responsibilities improve accountability',
        'Continuous improvement of processes'
    ]
})
with pd.ExcelWriter(OUTPUT_DIR / "final_assignment.xlsx", engine="openpyxl") as writer:
    save_to_excel(writer, task1_summary, "Task1_Summary")
    save_to_excel(writer, missing_details, "Task1_MissingDetails")
    save_to_excel(writer, recommendations, "Task1_Recommendations")
    save_to_excel(writer, top5_issues_freq, "Task2_Top5Issues_Freq")
    save_to_excel(writer, top5_issues_impact, "Task2_Top5Issues_Impact")
    save_to_excel(writer, cross_program_issues, "Task2_CrossProgram")
    save_to_excel(writer, resolution_analysis, "Task2_ResolutionAnalysis")
    save_to_excel(writer, single_fix, "Task2_SingleFix")
    save_to_excel(writer, weekly_trends, "Task2_WeeklyTrends")
    save_to_excel(writer, monthly_trends, "Task2_MonthlyTrends")
    save_to_excel(writer, quarterly_trends, "Task2_QuarterlyTrends")
    save_to_excel(writer, issue_distribution, "Task2_IssueDistribution")
    save_to_excel(writer, team_performance, "Task2_TeamPerformance")
    save_to_excel(writer, root_cause_summary, "Task3_RootCause")
    save_to_excel(writer, systemic_analysis, "Task3_SystemicAnalysis")
    save_to_excel(writer, process_mapping, "Task3_ProcessMapping")
    save_to_excel(writer, preventive_measures, "Task3_PreventiveMeasures")
    save_to_excel(writer, kpi_metrics, "Task3_KPIs")
    save_to_excel(writer, resource_recommendations, "Task3_Resources")
print("All Tasks completed. 19 sheets saved successfully.")