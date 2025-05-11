# streamlit_seating.py
import subprocess
import sys

# Ensure openpyxl is installed
subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl"], check=True)

import streamlit as st
import pandas as pd
# … rest of your imports and code …

import streamlit as st
import pandas as pd
import math
import io
from openpyxl import load_workbook

st.set_page_config(page_title="Exam Seating Arrangement", layout="wide")
st.title("📑 Exam Seating Arrangement Generator")

st.markdown("""
Upload your **input Excel** (with sheets:
- `in_timetable`
- `in_course_roll_mapping`
- `in_roll_name_mapping`
- `in_room_capacity`

Choose **buffer** & **density**, then download:
- Overall seating plan
- Seats-left report
- Morning routines (1 sheet per date)
- Evening routines (1 sheet per date)
""")

# 1) UPLOAD
uploaded = st.file_uploader("Upload Excel file", type="xlsx")
if not uploaded:
    st.info("Awaiting Excel file upload…")
    st.stop()

# 2) READ SHEETS
xls     = pd.ExcelFile(uploaded)
df_tt   = xls.parse("in_timetable")
df_cr   = xls.parse("in_course_roll_mapping")
df_rn   = xls.parse("in_roll_name_mapping")
df_room = xls.parse("in_room_capacity")

# 3) INPUTS
buffer  = st.number_input("Buffer seats per room", min_value=0, step=1, value=5)
density = st.radio("Density", ["Sparse", "Dense"])

# 4) PREP
df_cr.columns = df_cr.columns.str.strip()
df_cr['course_code'] = df_cr['course_code'].str.upper().str.strip()
df_cr['rollno']      = df_cr['rollno'].str.upper().str.strip()
df_rn.columns = df_rn.columns.str.strip()
df_rn['Roll'] = df_rn['Roll'].str.upper().str.strip()
name_dict = pd.Series(df_rn.Name.values, index=df_rn.Roll).to_dict()
course_to_rolls = df_cr.groupby('course_code')['rollno'].apply(lambda s: sorted(s.tolist())).to_dict()

rooms = []
for _,r in df_room.iterrows():
    rid   = str(r['Room No.']).strip()
    cap   = int(r['Exam Capacity'])
    block = str(r['Block']).strip()
    num   = int(rid) if block=='B1' else int(rid.split('-')[-1])
    rooms.append(dict(room=rid,capacity=cap,block=block,numeric=num))

def allocate_course(course, rolls, avail):
    N = len(rolls)
    allowed=[]
    for r in avail:
        eff = r['capacity']-buffer
        if eff<=0: continue
        use = math.floor(eff*0.5) if density=='Sparse' else eff
        if use>0:
            allowed.append(dict(**r,allowed=use))
    if sum(r['allowed'] for r in allowed) < N:
        return None
    b1=sorted([r for r in allowed if r['block']=='B1'], key=lambda x:x['numeric'])
    b2=sorted([r for r in allowed if r['block']=='B2'], key=lambda x:x['numeric'])
    # single block
    for pool in (b1,b2):
        if sum(r['allowed'] for r in pool)>=N:
            # find contiguous
            best=None
            for i in range(len(pool)):
                tot=0
                for j in range(i,len(pool)):
                    tot+=pool[j]['allowed']
                    if tot>=N:
                        L=j-i+1
                        if not best or L<best[0]:
                            best=(L,i,j)
                        break
            i,j=best[1],best[2]
            seg=pool[i:j+1]
            alloc=[]; rem=N; idx=0
            for rr in seg:
                t=min(rem, rr['allowed'])
                alloc.append((rr['room'], rolls[idx:idx+t]))
                idx+=t; rem-=t
            return alloc
    # split
    first,second = (b1,b2) if sum(r['allowed'] for r in b1)>=sum(r['allowed'] for r in b2) else (b2,b1)
    alloc=[]; rem=N; idx=0
    for rr in first:
        if rem<=0: break
        t=min(rem, rr['allowed'])
        alloc.append((rr['room'], rolls[idx:idx+t])); idx+=t; rem-=t
    for rr in second:
        if rem<=0: break
        t=min(rem, rr['allowed'])
        alloc.append((rr['room'], rolls[idx:idx+t])); idx+=t; rem-=t
    return None if rem>0 else alloc

# 5) ALLOCATE
overall=[]; seats_left=[]
morning_by_date={}; evening_by_date={}
for _,r in df_tt.iterrows():
    date = pd.to_datetime(r['Date']).strftime("%Y-%m-%d")
    morn = [] if pd.isna(r['Morning']) else [c.strip() for c in r['Morning'].split(';')]
    eve  = [] if pd.isna(r['Evening']) else [c.strip() for c in r['Evening'].split(';')]

    # Morning
    avail,overflow = rooms.copy(),[]
    for c in sorted(morn, key=lambda x:-len(course_to_rolls.get(x,[]))):
        rolls=course_to_rolls.get(c,[])
        if not rolls: continue
        alloc=allocate_course(c,rolls,avail)
        if not alloc: overflow.append(c)
        else:
            for rm,grp in alloc:
                avail=[x for x in avail if x['room']!=rm]
                overall.append({'Date':date,'Course':c,'Room':rm,'Count':len(grp),'Rolls':";".join(grp)})
                morning_by_date.setdefault(date,[]).append((c,rm,grp))

    # Evening
    avail=rooms.copy()
    for c in sorted(eve+overflow, key=lambda x:-len(course_to_rolls.get(x,[]))):
        rolls=course_to_rolls.get(c,[])
        if not rolls: continue
        alloc=allocate_course(c,rolls,avail)
        if not alloc:
            seats_left.append({'Date':date,'Course':c,'Unallocated':len(rolls)})
        else:
            for rm,grp in alloc:
                avail=[x for x in avail if x['room']!=rm]
                overall.append({'Date':date,'Course':c,'Room':rm,'Count':len(grp),'Rolls':";".join(grp)})
                evening_by_date.setdefault(date,[]).append((c,rm,grp))

# 6) DISPLAY
df_overall=pd.DataFrame(overall)
df_left  =pd.DataFrame(seats_left)
st.subheader("Overall Seating")
st.dataframe(df_overall)
st.subheader("Seats Left")
st.dataframe(df_left)

# 7) DOWNLOAD Overall & Seats Left
def to_xlsx(df,name):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine="openpyxl") as w:
        df.to_excel(w,sheet_name=name,index=False)
    buf.seek(0)
    return buf

c1,c2=st.columns(2)
c1.download_button("Download Overall",data=to_xlsx(df_overall,"Overall"),
                   file_name="op_overall.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
c2.download_button("Download Seats Left",data=to_xlsx(df_left,"SeatsLeft"),
                   file_name="op_seats_left.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# 8) BUILD Morning & Evening workbooks (one sheet per date)
# 8) BUILD Morning & Evening workbooks (one sheet per course-room, as in 2nd screenshot)
def build_course_workbook(allocs, session):
    """
    allocs: list of tuples (date, course, room, [rolls...])
    session: "Morning" or "Evening"
    """
    buf = io.BytesIO()
    # Create workbook and add each allocation as its own sheet
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for date, course, room, grp in allocs:
            sheet_name = f"{course} Room {room}"
            # Build a small DataFrame including header row as first row
            # We'll write the header row manually after saving to openpyxl
            df_sheet = pd.DataFrame(grp, columns=["Roll"]).assign(
                **{"Student Name": lambda d: d["Roll"].map(name_dict),
                   "Signature": ""}
            )
            # Write the table starting at row=2 (so row1 is free for our title)
            df_sheet.to_excel(writer, sheet_name=sheet_name,
                              index=False, startrow=1)
    buf.seek(0)

    # Now inject the title row in each sheet
    wb = load_workbook(buf)
    for date, course, room, grp in allocs:
        sheet_name = f"{course} Room {room}"
        ws = wb[sheet_name]
        title = f"Course: {course} | Room: {room} | Date: {date} | Session: {session}"
        ws.cell(row=1, column=1, value=title)
    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final

# Flatten morning_by_date into list of (date, course, room, grp)
morning_allocs = [
    (date, course, room, grp)
    for date, blocks in morning_by_date.items()
    for course, room, grp in blocks
]
evening_allocs = [
    (date, course, room, grp)
    for date, blocks in evening_by_date.items()
    for course, room, grp in blocks
]

buf_morn = build_course_workbook(morning_allocs, "Morning")
buf_eve  = build_course_workbook(evening_allocs,  "Evening")

# 9) DOWNLOAD Routines (unchanged)
c3, c4 = st.columns(2)
c3.download_button(
    "📥 Download Morning Routines",
    data=buf_morn,
    file_name="morning_routine.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
c4.download_button(
    "📥 Download Evening Routines",
    data=buf_eve,
    file_name="evening_routine.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


# 9) DOWNLOAD Routines
c3,c4=st.columns(2)
c3.download_button("Download Morning Routines",data=buf_morn,
                   file_name="morning_routine.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
c4.download_button("Download Evening Routines",data=buf_eve,
                   file_name="evening_routine.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.success("All outputs ready!")
