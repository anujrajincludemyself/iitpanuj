# #  streamlit_seating.py

# import streamlit as st
# import pandas as pd
# import math
# import io
# import zipfile
# from openpyxl import Workbook

# st.set_page_config(page_title="Exam Seating Arrangement", layout="wide")
# st.title("📑 Exam Seating Arrangement Generator")

# st.markdown("""
# Upload your **input Excel** (with sheets:
# - `in_timetable`
# - `in_course_roll_mapping`
# - `in_roll_name_mapping`
# - `in_room_capacity`

# Choose **buffer** & **density**, then download:
# - Overall seating plan
# - Seats-left report
# - **All schedules** as a single ZIP:
#   - overall_seating.xlsx
#   - seats_left.xlsx
#   - `<date>/morning_<date>.xlsx`
#   - `<date>/evening_<date>.xlsx`
# """)

# # 1) UPLOAD
# uploaded = st.file_uploader("Upload Excel file", type="xlsx")
# if not uploaded:
#     st.info("Awaiting Excel file upload…")
#     st.stop()

# # 2) READ SHEETS
# xls     = pd.ExcelFile(uploaded)
# df_tt   = xls.parse("in_timetable")
# df_cr   = xls.parse("in_course_roll_mapping")
# df_rn   = xls.parse("in_roll_name_mapping")
# df_room = xls.parse("in_room_capacity")

# # 3) INPUTS
# buffer  = st.number_input("Buffer seats per room", min_value=0, step=1, value=5)
# density = st.radio("Density", ["Sparse", "Dense"])

# # 4) PREP
# df_cr.columns = df_cr.columns.str.strip()
# df_cr['course_code'] = df_cr['course_code'].str.upper().str.strip()
# df_cr['rollno']      = df_cr['rollno'].str.upper().str.strip()
# df_rn.columns = df_rn.columns.str.strip()
# df_rn['Roll'] = df_rn['Roll'].str.upper().str.strip()
# name_dict = pd.Series(df_rn.Name.values, index=df_rn.Roll).to_dict()
# course_to_rolls = (
#     df_cr.groupby('course_code')['rollno']
#          .apply(lambda s: sorted(s.tolist()))
#          .to_dict()
# )

# rooms = []
# for _, r in df_room.iterrows():
#     rid   = str(r['Room No.']).strip()
#     cap   = int(r['Exam Capacity'])
#     block = str(r['Block']).strip()
#     num   = int(rid) if block=='B1' else int(rid.split('-')[-1])
#     rooms.append(dict(room=rid, capacity=cap, block=block, numeric=num))

# def allocate_course(course, rolls, avail):
#     N = len(rolls)
#     allowed=[]
#     for r in avail:
#         eff = r['capacity'] - buffer
#         if eff<=0: continue
#         use = math.floor(eff*0.5) if density=='Sparse' else eff
#         if use>0:
#             allowed.append(dict(**r, allowed=use))
#     if sum(r['allowed'] for r in allowed) < N:
#         return None
#     b1 = sorted([r for r in allowed if r['block']=='B1'], key=lambda x:x['numeric'])
#     b2 = sorted([r for r in allowed if r['block']=='B2'], key=lambda x:x['numeric'])
#     for pool in (b1,b2):
#         if sum(r['allowed'] for r in pool) >= N:
#             best=None
#             for i in range(len(pool)):
#                 tot=0
#                 for j in range(i,len(pool)):
#                     tot+=pool[j]['allowed']
#                     if tot>=N:
#                         L=j-i+1
#                         if not best or L<best[0]:
#                             best=(L,i,j)
#                         break
#             i,j = best[1], best[2]
#             seg=pool[i:j+1]
#             alloc, rem, idx = [], N, 0
#             for rr in seg:
#                 t = min(rem, rr['allowed'])
#                 alloc.append((rr['room'], rolls[idx:idx+t]))
#                 idx+=t; rem-=t
#             return alloc
#     first,second = (b1,b2) if sum(r['allowed'] for r in b1)>=sum(r['allowed'] for r in b2) else (b2,b1)
#     alloc, rem, idx = [], N, 0
#     for rr in first:
#         if rem<=0: break
#         t=min(rem,rr['allowed'])
#         alloc.append((rr['room'], rolls[idx:idx+t])); idx+=t; rem-=t
#     for rr in second:
#         if rem<=0: break
#         t=min(rem,rr['allowed'])
#         alloc.append((rr['room'], rolls[idx:idx+t])); idx+=t; rem-=t
#     return None if rem>0 else alloc

# # 5) ALLOCATE
# overall=[]; seats_left=[]
# morning_by_date={}; evening_by_date={}

# for _, row in df_tt.iterrows():
#     date = pd.to_datetime(row['Date']).strftime("%Y-%m-%d")
#     morn = [] if pd.isna(row['Morning']) else [c.strip() for c in row['Morning'].split(';')]
#     eve  = [] if pd.isna(row['Evening']) else [c.strip() for c in row['Evening'].split(';')]

#     avail, overflow = rooms.copy(), []
#     for c in sorted(morn, key=lambda x:-len(course_to_rolls.get(x,[]))):
#         rolls=course_to_rolls.get(c,[])
#         if not rolls: continue
#         alloc=allocate_course(c,rolls,avail)
#         if not alloc: overflow.append(c)
#         else:
#             for rm,grp in alloc:
#                 avail=[x for x in avail if x['room']!=rm]
#                 overall.append({'Date':date,'Course':c,'Room':rm,'Count':len(grp),'Rolls':";".join(grp)})
#                 morning_by_date.setdefault(date,[]).append((c,rm,grp))

#     avail=rooms.copy()
#     for c in sorted(eve+overflow, key=lambda x:-len(course_to_rolls.get(x,[]))):
#         rolls=course_to_rolls.get(c,[])
#         if not rolls: continue
#         alloc=allocate_course(c,rolls,avail)
#         if not alloc:
#             seats_left.append({'Date':date,'Course':c,'Unallocated':len(rolls)})
#         else:
#             for rm,grp in alloc:
#                 avail=[x for x in avail if x['room']!=rm]
#                 overall.append({'Date':date,'Course':c,'Room':rm,'Count':len(grp),'Rolls':";".join(grp)})
#                 evening_by_date.setdefault(date,[]).append((c,rm,grp))

# # 6) DISPLAY
# df_overall = pd.DataFrame(overall)
# df_left    = pd.DataFrame(seats_left)

# st.subheader("Overall Seating Plan")
# st.dataframe(df_overall)
# st.subheader("Seats Left Report")
# st.dataframe(df_left)

# def buf_for_df(df, sheet):
#     buf=io.BytesIO()
#     with pd.ExcelWriter(buf, engine="openpyxl") as w:
#         df.to_excel(w, sheet_name=sheet, index=False)
#     buf.seek(0)
#     return buf

# # Create ZIP
# zip_buf = io.BytesIO()
# with zipfile.ZipFile(zip_buf, "w") as z:
#     z.writestr("overall_seating.xlsx", buf_for_df(df_overall,"Overall").getvalue())
#     z.writestr("seats_left.xlsx", buf_for_df(df_left,"SeatsLeft").getvalue())

#     for date, blocks in morning_by_date.items():
#         wb = Workbook()
#         wb.remove(wb.active)
#         for course, room, grp in blocks:
#             sheet = wb.create_sheet(title=f"{course} Room {room}")
#             sheet.append([f"Course: {course} | Room: {room} | Date: {date} | Session: Morning"])
#             sheet.append(["Roll","Student Name","Signature"])
#             for rn in grp:
#                 sheet.append([rn, name_dict.get(rn,"Unknown Name"), ""])
#         out = io.BytesIO(); wb.save(out)
#         z.writestr(f"{date}/morning_{date}.xlsx", out.getvalue())

#         evening_blocks = evening_by_date.get(date, [])
#         if evening_blocks:
#             wb = Workbook()
#             wb.remove(wb.active)
#             for course, room, grp in evening_blocks:
#                 sheet = wb.create_sheet(title=f"{course} Room {room}")
#                 sheet.append([f"Course: {course} | Room: {room} | Date: {date} | Session: Evening"])
#                 sheet.append(["Roll","Student Name","Signature"])
#                 for rn in grp:
#                     sheet.append([rn, name_dict.get(rn,"Unknown Name"), ""])
#             out = io.BytesIO(); wb.save(out)
#             z.writestr(f"{date}/evening_{date}.xlsx", out.getvalue())

# zip_buf.seek(0)

# st.download_button(
#     "📥 Download All Schedules (ZIP)",
#     data=zip_buf,
#     file_name="schedules.zip",
#     mime="application/zip"
# )

# st.success("✅ schedules.zip is ready — click to download!")


import streamlit as st
import pandas as pd
import math
import io
import zipfile
from openpyxl import Workbook

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
- **All schedules** as a single ZIP:
  - overall_seating.xlsx
  - seats_left.xlsx
  - `<date>/morning_<date>.xlsx`
  - `<date>/evening_<date>.xlsx`
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
course_to_rolls = (
    df_cr.groupby('course_code')['rollno']
         .apply(lambda s: sorted(s.tolist()))
         .to_dict()
)

rooms = []
for _, r in df_room.iterrows():
    rid   = str(r['Room No.']).strip()
    cap   = int(r['Exam Capacity'])
    block = str(r['Block']).strip()
    num   = int(rid) if block=='B1' else int(rid.split('-')[-1])
    rooms.append(dict(room=rid, capacity=cap, block=block, numeric=num))

def allocate_course(course, rolls, avail):
    N = len(rolls)
    allowed=[]
    for r in avail:
        eff = r['capacity'] - buffer
        if eff<=0: continue
        use = math.floor(eff*0.5) if density=='Sparse' else eff
        if use>0:
            allowed.append(dict(**r, allowed=use))
    if sum(r['allowed'] for r in allowed) < N:
        return None
    b1 = sorted([r for r in allowed if r['block']=='B1'], key=lambda x:x['numeric'])
    b2 = sorted([r for r in allowed if r['block']=='B2'], key=lambda x:x['numeric'])
    for pool in (b1,b2):
        if sum(r['allowed'] for r in pool) >= N:
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
            i,j = best[1], best[2]
            seg=pool[i:j+1]
            alloc, rem, idx = [], N, 0
            for rr in seg:
                t = min(rem, rr['allowed'])
                alloc.append((rr['room'], rolls[idx:idx+t]))
                idx+=t; rem-=t
            return alloc
    first,second = (b1,b2) if sum(r['allowed'] for r in b1)>=sum(r['allowed'] for r in b2) else (b2,b1)
    alloc, rem, idx = [], N, 0
    for rr in first:
        if rem<=0: break
        t=min(rem,rr['allowed'])
        alloc.append((rr['room'], rolls[idx:idx+t])); idx+=t; rem-=t
    for rr in second:
        if rem<=0: break
        t=min(rem,rr['allowed'])
        alloc.append((rr['room'], rolls[idx:idx+t])); idx+=t; rem-=t
    return None if rem>0 else alloc

# 5) ALLOCATE
overall=[]; seats_left=[]
morning_by_date={}; evening_by_date={}

for _, row in df_tt.iterrows():
    date = pd.to_datetime(row['Date']).strftime("%d-%m-%Y")  # <-- date format fixed here
    morn = [] if pd.isna(row['Morning']) else [c.strip() for c in row['Morning'].split(';')]
    eve  = [] if pd.isna(row['Evening']) else [c.strip() for c in row['Evening'].split(';')]

    avail, overflow = rooms.copy(), []
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
df_overall = pd.DataFrame(overall)
df_left    = pd.DataFrame(seats_left)

st.subheader("Overall Seating Plan")
st.dataframe(df_overall)
st.subheader("Seats Left Report")
st.dataframe(df_left)

def buf_for_df(df, sheet):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=sheet, index=False)
    buf.seek(0)
    return buf

# Create ZIP
zip_buf = io.BytesIO()
with zipfile.ZipFile(zip_buf, "w") as z:
    z.writestr("overall_seating.xlsx", buf_for_df(df_overall,"Overall").getvalue())
    z.writestr("seats_left.xlsx", buf_for_df(df_left,"SeatsLeft").getvalue())

    for date, blocks in morning_by_date.items():
        wb = Workbook()
        wb.remove(wb.active)
        for course, room, grp in blocks:
            sheet = wb.create_sheet(title=f"{course} Room {room}")
            sheet.append([f"Course: {course} | Room: {room} | Date: {date} | Session: Morning"])
            sheet.append(["Roll","Student Name","Signature"])
            for rn in grp:
                sheet.append([rn, name_dict.get(rn,"Unknown Name"), ""])
        out = io.BytesIO(); wb.save(out)
        z.writestr(f"{date}/morning_{date}.xlsx", out.getvalue())

        evening_blocks = evening_by_date.get(date, [])
        if evening_blocks:
            wb = Workbook()
            wb.remove(wb.active)
            for course, room, grp in evening_blocks:
                sheet = wb.create_sheet(title=f"{course} Room {room}")
                sheet.append([f"Course: {course} | Room: {room} | Date: {date} | Session: Evening"])
                sheet.append(["Roll","Student Name","Signature"])
                for rn in grp:
                    sheet.append([rn, name_dict.get(rn,"Unknown Name"), ""])
            out = io.BytesIO(); wb.save(out)
            z.writestr(f"{date}/evening_{date}.xlsx", out.getvalue())

zip_buf.seek(0)

st.download_button(
    "📥 Download All Schedules (ZIP)",
    data=zip_buf,
    file_name="schedules.zip",
    mime="application/zip"
)

st.success("✅ schedules.zip is ready — click to download!")
