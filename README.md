=> Features 

=>   Upload input Excel file with:

(1)  in_timetable

(2) in_course_roll_mapping

(3) in_roll_name_mapping

(4) in_room_capacity

=> Select:

Buffer seats per room

Seating Density (Sparse / Dense)

📈 Automatically:

Allocates students room-wise and session-wise (Morning / Evening)

Ensures capacity and buffer constraints

Attempts block-wise (B1 first, then B2) allocations

Generates attendance sheets per room-session-date

Provides a summary of unallocated students (if any)

📦 Exports:

overall_seating.xlsx — complete seating plan

seats_left.xlsx — unallocated students (if any)

Per-date, per-session attendance sheets inside a downloadable ZIP

📁 Example Output
python
Copy
Edit
schedules.zip
├── 12_06_2025/
│   ├── morning/
│   │   ├── 12_06_2025_MA101_101_morning.xlsx
│   │   └── ...
│   └── evening/
│       ├── 12_06_2025_CS102_203_evening.xlsx
│       └── ...
├── overall_seating.xlsx
└── seats_left.xlsx


🛠️ Technologies Used
Python

Streamlit

Pandas

OpenPyXL

Zipfile




https://iitpanuj-avnf47wshc3dokfmbuopbk.streamlit.app/ ( here you can use the app)
