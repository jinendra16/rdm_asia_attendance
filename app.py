import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime, timedelta
import re
import io

# --- PAGE CONFIG ---
st.set_page_config(page_title="RDM Timesheet Auditor", page_icon="📊")

def parse_date_manual(user_input):
    try:
        clean_input = re.split(r'[_ ]', user_input)
        day = int(clean_input[0])
        month_str = clean_input[1].capitalize()[:3]
        current_year = 2026 
        month_map = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'Jun': 6,
                     'Jul': 7, 'Aug': 8, 'Sep': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12}
        return datetime(current_year, month_map.get(month_str, 1), day)
    except:
        return None

def clean_name(name):
    if pd.isna(name): return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(name).upper())

def process_data(ts_file, att_file, start_date):
    # Date Setup
    end_date = start_date + timedelta(days=6)
    date_range = [(start_date + timedelta(days=i)).date() for i in range(7)]
    dynamic_sheet_name = f"{start_date.day} {start_date.strftime('%b')} - {end_date.day} {end_date.strftime('%b')}"

    # Load Files
    if ts_file.name.endswith('.csv'):
        df = pd.read_csv(ts_file)
    else:
        df = pd.read_excel(ts_file)
    
    df['DateTime'] = pd.to_datetime(
        df['Date Time'].str.extract(r'(\d{4}-\d{2}-\d{2})')[0] + ' ' + 
        df['Date Time'].str.extract(r'(\d{2}:\d{2})')[0]
    )
    
    df['WorkDate'] = df.apply(lambda r: (r['DateTime'] - timedelta(hours=8)).date(), axis=1)
    df['TimeOnly'] = df['DateTime'].dt.time
    df['CleanName'] = df['Name'].apply(clean_name)

    try:
        df_attendance = pd.read_excel(att_file, sheet_name=dynamic_sheet_name, header=None)
    except:
        df_attendance = pd.read_excel(att_file, sheet_name=0, header=None)
    
    raw_employee_list = [df_attendance.iloc[idx, 1] for idx in range(2, min(100, len(df_attendance))) 
                         if pd.notna(df_attendance.iloc[idx, 1])]
    attendance_map = {clean_name(name): name for name in raw_employee_list}
    
    output_data = []
    exception_logs = []

    for cleaned_name, original_name in attendance_map.items():
        row_times = []
        emp_records = df[df['CleanName'] == cleaned_name].sort_values('DateTime')

        for work_date in date_range:
            day_logs = emp_records[emp_records['WorkDate'] == work_date]
            login_time, logout_time = None, None
            
            if not day_logs.empty:
                # Login Logic
                start_work_logs = day_logs[day_logs['Type'] == 'Start Work']
                site_in_logs = day_logs[day_logs['Type'] == 'Site In']
                if not start_work_logs.empty: login_time = start_work_logs.iloc[0]['TimeOnly']
                elif not site_in_logs.empty: login_time = site_in_logs.iloc[0]['TimeOnly']
                else: login_time = day_logs.iloc[0]['TimeOnly']

                # Logout Logic
                end_work_logs = day_logs[day_logs['Type'] == 'End Work']
                site_out_logs = day_logs[day_logs['Type'] == 'Site Out']
                if not end_work_logs.empty: logout_time = end_work_logs.iloc[-1]['TimeOnly']
                elif not site_out_logs.empty: logout_time = site_out_logs.iloc[-1]['TimeOnly']
                else: logout_time = "NO LOGOUT"

                if len(day_logs) == 1 and logout_time != "NO LOGOUT":
                    if any(x in str(day_logs.iloc[0]['Type']) for x in ['Start Work', 'Site In']):
                        logout_time = "NO LOGOUT"

                if logout_time == "NO LOGOUT":
                    exception_logs.append({'Name': original_name, 'Date': work_date, 'Time': login_time, 'Reason': 'Missing Logout'})
                elif "Site In" in day_logs['Type'].values and "Site Out" not in day_logs['Type'].values:
                    exception_logs.append({'Name': original_name, 'Date': work_date, 'Time': logout_time, 'Reason': 'Missing Site Out'})

            row_times.extend([login_time, logout_time])
        output_data.append(row_times)

    return pd.DataFrame(output_data), pd.DataFrame(exception_logs), date_range, raw_employee_list

# --- UI DESIGN ---
st.title("📊 Timesheet Audit Web Portal")
st.markdown("Upload your files below to generate the audit report.")

col1, col2 = st.columns(2)

with col1:
    ts_file = st.file_uploader("Upload Timesheet Detail (Excel/CSV)", type=['xlsx', 'csv'])
with col2:
    att_file = st.file_uploader("Upload Attendance_New.xlsx", type=['xlsx'])

start_date_str = st.text_input("Enter Start Date (e.g., 23_Jan)", "")

if st.button("🚀 Generate Audit Report"):
    if ts_file and att_file and start_date_str:
        start_date = parse_date_manual(start_date_str)
        if start_date:
            with st.spinner('Processing...'):
                df_out, df_ex, d_range, names = process_data(ts_file, att_file, start_date)
                
                # Show Preview
                st.success("Analysis Complete!")
                st.subheader("Preview (Summary)")
                st.dataframe(df_out)
                
                # Prepare Excel for download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # You can use your save_to_excel logic here or simply:
                    df_out.to_excel(writer, index=False, sheet_name="Summary")
                    df_ex.to_excel(writer, index=False, sheet_name="Exceptions")
                
                st.download_button(
                    label="📥 Download Audit Report",
                    data=output.getvalue(),
                    file_name=f"Audit_Report_{start_date.strftime('%d_%b')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.error("Invalid Date Format. Please use '23_Jan'")
    else:
        st.warning("Please upload both files and enter a start date.")