import os
import pandas as pd
from datetime import datetime
import streamlit as st

# Load the input data
def load_data(file_path):
    data = pd.read_excel(file_path, sheet_name=None, header=None)

    # Process ip_1
    ip_1_raw = data['ip_1']
    ip_1_raw.columns = ip_1_raw.iloc[1].str.strip().str.lower()
    ip_1 = ip_1_raw[2:].reset_index(drop=True)  # Skip the first two rows
    ip_1.columns = ip_1.columns.str.strip().str.lower()

    # Process ip_2 (Exam timetable)
    ip_2_raw = data['ip_2']
    ip_2 = ip_2_raw.rename(columns={0: 'date', 1: 'day', 2: 'morning', 3: 'evening'})
    ip_2 = ip_2[1:].reset_index(drop=True)  # Remove header row

    # Process ip_3 (Room capacity and block details)
    ip_3 = data['ip_3']
    ip_3.columns = ['room_no', 'exam_capacity', 'block']
    ip_3['exam_capacity'] = pd.to_numeric(ip_3['exam_capacity'], errors='coerce')

    # Process ip_4 (Roll and Name mapping)
    ip_4 = data['ip_4']
    ip_4.columns = ['roll_no', 'student_name']

    return ip_1, ip_2, ip_3, ip_4

# Allocate rooms for exams
def allocate_rooms(ip_1, ip_3, buffer=5, mode='dense'):
    ip_1['course_code'] = ip_1['course_code'].str.strip()
    course_counts = ip_1['course_code'].value_counts().to_dict()
    sorted_courses = sorted(course_counts.keys(), key=lambda x: course_counts[x], reverse=True)

    ip_3['block'] = ip_3['block'].astype(str).str.strip()
    ip_3 = ip_3.sort_values(by=['block', 'room_no'], ascending=[True, True])

    room_allocations = []

    for _, room in ip_3.iterrows():
        room_name = room['room_no']
        room_capacity = room['exam_capacity'] - buffer
        if pd.isna(room_capacity) or room_capacity <= 0:
            continue

        room_fill = 0
        for course in sorted_courses[:]:
            course_students = ip_1[ip_1['course_code'] == course]

            if mode == 'dense' and len(course_students) <= room_capacity - room_fill:
                room_allocations.append({
                    'room': room_name,
                    'course': course,
                    'students': len(course_students),
                    'vacant_seats': room_capacity - room_fill - len(course_students),
                    'block': room['block']
                })
                room_fill += len(course_students)
                sorted_courses.remove(course)

    return pd.DataFrame(room_allocations)

# Generate seating plan
def generate_seating_plan(ip_1, ip_2, ip_4, room_allocations):
    seating_plan = []
    for _, allocation in room_allocations.iterrows():
        course_students = ip_1[ip_1['course_code'] == allocation['course']]
        merged = pd.merge(course_students, ip_4, left_on='rollno', right_on='roll_no', how='left')
        room = allocation['room']
        roll_list = ";".join(merged['roll_no'].tolist())

        date, day, session = None, None, None
        for _, row in ip_2.iterrows():
            if allocation['course'] in str(row['morning']):
                date, day, session = row['date'], row['day'], 'morning'
                break
            elif allocation['course'] in str(row['evening']):
                date, day, session = row['date'], row['day'], 'evening'
                break

        if date is not None:
            seating_plan.append({
                'Date': pd.to_datetime(date).strftime("%d/%m/%Y"),
                'Day': day,
                'course_code': allocation['course'],
                'Room': room,
                'Allocated_students_count': len(merged),
                'Roll_list': roll_list,
                'Session': session
            })

    return pd.DataFrame(seating_plan)

# Generate attendance sheet
def generate_attendance_sheet(seating_plan, ip_4):
    attendance_sheets = []

    for _, row in seating_plan.iterrows():
        room = row['Room']
        roll_list = row['Roll_list'].split(";")
        course_code = row['course_code']
        date = row['Date']
        session = row['Session']

        attendance_data = ip_4[ip_4['roll_no'].isin(roll_list)].copy()
        attendance_data['signature'] = ""

        # Add blank rows for invigilators and TAs
        blank_rows = pd.DataFrame({
            'roll_no': [""] * 5,
            'student_name': [""] * 5,
            'signature': [""] * 5
        })
        attendance_data = pd.concat([attendance_data, blank_rows], ignore_index=True)

        attendance_sheets.append({
            'Date': date,
            'Room': room,
            'Session': session,
            'Course_Code': course_code,
            'Attendance_Data': attendance_data
        })

    return attendance_sheets

# Streamlit application
def main():
    st.title("Exam Room Allocation and Seating Plan")
    
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    if uploaded_file is not None:
        buffer = st.number_input("Buffer Seats Per Room", min_value=0, value=5)
        mode = st.selectbox("Room Allocation Mode", options=["dense", "sparse"], index=0)

        ip_1, ip_2, ip_3, ip_4 = load_data(uploaded_file)

        st.subheader("Room Allocations")
        room_allocations = allocate_rooms(ip_1, ip_3, buffer=buffer, mode=mode)
        st.dataframe(room_allocations)

        st.subheader("Seating Plan")
        seating_plan = generate_seating_plan(ip_1, ip_2, ip_4, room_allocations)
        st.dataframe(seating_plan)

        st.subheader("Attendance Sheets")
        attendance_sheets = generate_attendance_sheet(seating_plan, ip_4)

        for sheet in attendance_sheets:
            st.write(f"Room: {sheet['Room']}, Date: {sheet['Date']}, Session: {sheet['Session']}")
            st.dataframe(sheet['Attendance_Data'])

        # Download links for CSV and Excel
        seating_csv = seating_plan.to_csv(index=False).encode('utf-8')
        st.download_button(label="Download Seating Plan (CSV)", data=seating_csv, file_name="seating_plan.csv", mime="text/csv")

        room_csv = room_allocations.to_csv(index=False).encode('utf-8')
        st.download_button(label="Download Room Allocations (CSV)", data=room_csv, file_name="room_allocations.csv", mime="text/csv")

if __name__ == "__main__":
    main()