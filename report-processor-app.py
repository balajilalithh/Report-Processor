import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
import io
import matplotlib.pyplot as plt
import plotly.express as px

# Helper functions
def get_next_friday(start_date):
    days_ahead = 4 - start_date.weekday()  # Friday is the 4th day of the week
    if days_ahead <= 0:
        days_ahead += 7
    return start_date + timedelta(days_ahead)

def add_business_days(start_date, days):
    current_date = start_date
    while days > 0:
        current_date += timedelta(1)
        if current_date.weekday() < 5:  # Mon-Fri are considered business days
            days -= 1
    return current_date

# Streamlit app setup
st.title("Ticket Report Generator")
uploaded_file = st.file_uploader("Choose an Excel or CSV file", type=["xlsx", "xls", "csv"])

# File reading
if uploaded_file is not None:
    file_extension = uploaded_file.name.split('.')[-1]
    if file_extension == 'xlsx' or file_extension == 'xls':
        df = pd.read_excel(uploaded_file)
    elif file_extension == 'csv':
        df = pd.read_csv(uploaded_file)

    # Date conversion
    df['Created Time (Ticket)'] = pd.to_datetime(df['Created Time (Ticket)'], format='%d %b %Y %I:%M %p')
    df['Ticket Closed Time'] = pd.to_datetime(df['Ticket Closed Time'], format='%d %b %Y %I:%M %p', errors='coerce')

    # Exclude 'ElastAlert' tickets
    df_non_elastalert = df[~df['Subject'].str.contains("ElastAlert", case=False)]

    # Expected TAT definition
    expected_tat = {'P1': 1, 'P2': 3, 'P3': 7, 'P4': 30}

    # Summary Report
    st.header("Summary Report")
    summary_report_data = []
    for priority in ['P1', 'P2', 'P3', 'P4']:
        priority_df = df_non_elastalert[df_non_elastalert['Priority (Ticket)'].str.contains(priority, case=False)]
        not_an_issue_df = priority_df[priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])]
        not_an_issue = len(not_an_issue_df[not_an_issue_df['Category (Ticket)'].str.lower().isin(['query', 'access request'])])
        closed_tickets_count = len(priority_df[priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])])
        
        if closed_tickets_count > 0:
            closed_tickets_df = priority_df[priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])]
            closed_tickets_df['TAT'] = (closed_tickets_df['Ticket Closed Time'] - closed_tickets_df['Created Time (Ticket)']).dt.days + 1
            actual_tat = closed_tickets_df['TAT'].mean()
        else:
            actual_tat = 'NA'
        
        summary_report_data.append({
            'Priority': priority,
            'Not an Issue': not_an_issue,
            'Closed Tickets': closed_tickets_count,
            'Expected TAT': expected_tat[priority],
            'Actual TAT': round(actual_tat, 2) if actual_tat != 'NA' else actual_tat
        })
    
    summary_report_df = pd.DataFrame(summary_report_data)
    st.dataframe(summary_report_df)

    # Provide download link for the summary report
    summary_excel = io.BytesIO()
    with pd.ExcelWriter(summary_excel, engine='xlsxwriter') as writer:
        summary_report_df.to_excel(writer, index=False)
    summary_excel.seek(0)
    st.download_button(
        label="Download Summary Report",
        data=summary_excel,
        file_name="summary_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Detailed Report
    st.header("Detailed Report")
    detailed_report_data = []
    for priority in ['P1', 'P2', 'P3', 'P4']:
        priority_df = df_non_elastalert[df_non_elastalert['Priority (Ticket)'].str.contains(priority, case=False)]
        tickets_raised = len(priority_df)
        not_an_issue_df = priority_df[priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])]
        not_an_issue = len(not_an_issue_df)
        bugs = len(priority_df[priority_df['Category (Ticket)'].str.lower() == 'bug'])
        closed_tickets_count = len(priority_df[priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])])
        
        if closed_tickets_count > 0:
            closed_tickets_df = priority_df[priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])]
            closed_tickets_df['TAT'] = (closed_tickets_df['Ticket Closed Time'] - closed_tickets_df['Created Time (Ticket)']).dt.days + 1
            actual_tat = closed_tickets_df['TAT'].mean()
        else:
            actual_tat = 'NA'

        pending_tickets_count = len(priority_df[~priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])])
        if pending_tickets_count > 0:
            pending_tickets_df = priority_df[~priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])]
            pending_tickets_df['Target ETA'] = pending_tickets_df['Created Time (Ticket)'].apply(get_next_friday)
        else:
            pending_tickets_df = None
        
        detailed_report_data.append({
            'Ticket Priority': priority,
            'Tickets raised': tickets_raised,
            'Not an Issue': not_an_issue,
            'Bugs': bugs,
            'Tickets Closed': closed_tickets_count,
            'Expected TAT': expected_tat[priority],
            'Actual TAT': round(actual_tat, 2) if actual_tat != 'NA' else actual_tat,
            'Pending Tickets': pending_tickets_count,
            'Target ETA': pending_tickets_df['Target ETA'].max().strftime('%d %b %Y') if pending_tickets_df is not None else 'NA'
        })
    
    detailed_report_df = pd.DataFrame(detailed_report_data)
    st.dataframe(detailed_report_df)

    # Provide download link for the detailed report
    detailed_excel = io.BytesIO()
    with pd.ExcelWriter(detailed_excel, engine='xlsxwriter') as writer:
        detailed_report_df.to_excel(writer, index=False)
    detailed_excel.seek(0)
    st.download_button(
        label="Download Detailed Report",
        data=detailed_excel,
        file_name="detailed_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Alerts Report
    st.header("Alerts Report")
    alerts_df = df[df['Subject'].str.contains("ElastAlert", case=False)]
    alerts_report_data = []
    for priority in ['P1', 'P2', 'P3', 'P4']:
        priority_df = alerts_df[alerts_df['Priority (Ticket)'].str.contains(priority, case=False)]
        not_an_issue_df = priority_df[priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])]
        not_an_issue = len(not_an_issue_df[not_an_issue_df['Category (Ticket)'].str.lower().isin(['query', 'access request'])])
        closed_tickets_count = len(priority_df[priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])])
        
        if closed_tickets_count > 0:
            closed_tickets_df = priority_df[priority_df['Status (Ticket)'].str.lower().isin(['closed', 'duplicate'])]
            closed_tickets_df['TAT'] = (closed_tickets_df['Ticket Closed Time'] - closed_tickets_df['Created Time (Ticket)']).dt.days + 1
            actual_tat = closed_tickets_df['TAT'].mean()
        else:
            actual_tat = 'NA'
        
        alerts_report_data.append({
            'Priority': priority,
            'Not an Issue': not_an_issue,
            'Closed Tickets': closed_tickets_count,
            'Expected TAT': expected_tat[priority],
            'Actual TAT': round(actual_tat, 2) if actual_tat != 'NA' else actual_tat
        })
    
    alerts_report_df = pd.DataFrame(alerts_report_data)
    st.dataframe(alerts_report_df)

    # Provide download link for the alerts report
    alerts_excel = io.BytesIO()
    with pd.ExcelWriter(alerts_excel, engine='xlsxwriter') as writer:
        alerts_report_df.to_excel(writer, index=False)
    alerts_excel.seek(0)
    st.download_button(
        label="Download Alerts Report",
        data=alerts_excel,
        file_name="alerts_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Graphs
    st.header("Graphs")

    # Summary Report Graph
    fig_summary = px.bar(summary_report_df, x='Priority', y='Closed Tickets', title='Closed Tickets by Priority')
    st.plotly_chart(fig_summary)

    # Detailed Report Graph
    fig_detailed = px.bar(detailed_report_df, x='Ticket Priority', y='Tickets Closed', title='Tickets Closed by Priority')
    st.plotly_chart(fig_detailed)

    # Alerts Report Graph
    fig_alerts = px.bar(alerts_report_df, x='Priority', y='Closed Tickets', title='Closed Tickets by Priority (ElastAlert)')
    st.plotly_chart(fig_alerts)
