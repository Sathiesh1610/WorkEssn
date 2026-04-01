import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime, timedelta

# =========================================================
# 🖥️ UI SETUP - File selection dialogs
# =========================================================

# Hide the main tkinter window
root = tk.Tk()
root.withdraw()

# --- Select Excel file ---
excel_path = filedialog.askopenfilename(
    title="Select Shift Roster Excel File",
    filetypes=[("Excel Files", "*.xlsx *.xls")]
)

if not excel_path:
    print("No Excel file selected. Exiting...")
    exit()

# --- Select destination folder for ICS file ---
save_folder = filedialog.askdirectory(
    title="Select Folder to Save ICS File"
)

if not save_folder:
    print("No folder selected. Exiting...")
    exit()

# =========================================================
# 📁 FILE NAMING - Auto format: ShiftRooster_MMMYY.ics
# =========================================================

current_month = datetime.now().strftime("%b%y")  # Example: Apr26
filename = f"ShiftRooster_{current_month}.ics"
output_path = os.path.join(save_folder, filename)

# =========================================================
# ⚙️ HELPER FUNCTIONS & SHIFT CONFIG
# =========================================================

# Format datetime into ICS required format
def dtfmt(dt):
    return dt.strftime("%Y%m%dT%H%M%S")

# Shift timings (start, end)
shift_time = {
    "S1": ("06:00", "15:30"),
    "G": ("09:00", "18:30"),
    "S2": ("14:00", "23:30"),
    "EVE": ("17:00", "02:30"),
    "S3": ("22:00", "07:30")
}

# Shifts considered as working
WORKING_SHIFTS = {"S1", "G", "S2", "EVE", "S3"}

"""
# =========================================================
# ⚠️ TEMP DATA (TO BE REPLACED WITH EXCEL IN NEXT STEP)
# =========================================================

# NOTE: These must exist for the script to run
# You will REMOVE this section once Excel integration is added

dates = ["01-Apr", "02-Apr"]  # Example dates
names = ["Sathiesh M", "John"]

shifts = {
    "Sathiesh M": ["S1", "OFF"],
    "John": ["S1", "S2"]
}
"""
# =========================================================
# 📊 READ EXCEL FILE
# =========================================================

df = pd.read_excel(excel_path)

# First column = Names
names = df.iloc[:, 0].dropna().tolist()

# Remaining columns = Dates
date_columns = df.columns[1:]

# Convert column headers to required format (e.g., 01-Apr)
dates = [pd.to_datetime(col).strftime("%d-%b") for col in date_columns]

# Build shifts dictionary
shifts = {}

for i, name in enumerate(names):
    shifts[name] = []

    for col in date_columns:
        value = df.iloc[i][col]

        if pd.isna(value):
            shifts[name].append("OFF")
        else:
            shifts[name].append(str(value).strip())

# =========================================================
# 📅 ICS FILE GENERATION
# =========================================================

lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Shift Calendar//Sathiesh M//EN"
]

# Loop through each day
for i, d in enumerate(dates):

    # Get Sathiesh's shift for the day
    sath = shifts["Sathiesh M"][i]

    # Convert date string to datetime object
    ymd = datetime.strptime(f"{d}-2026", "%d-%b-%Y")

    lines.append("BEGIN:VEVENT")

    # =====================================================
    # 🟢 WORKING DAY EVENT
    # =====================================================
    if sath in shift_time:

        st, et = shift_time[sath]

        start_dt = datetime.strptime(f"{ymd:%Y-%m-%d} {st}", "%Y-%m-%d %H:%M")
        end_dt = datetime.strptime(f"{ymd:%Y-%m-%d} {et}", "%Y-%m-%d %H:%M")

        # Handle overnight shifts
        if sath in ("S3", "EVE"):
            end_dt += timedelta(days=1)

        # Add time details
        lines.append(f"DTSTART;TZID=Asia/Kolkata:{dtfmt(start_dt)}")
        lines.append(f"DTEND;TZID=Asia/Kolkata:{dtfmt(end_dt)}")
        lines.append(f"SUMMARY:{sath} ({st}–{et})")

        # =================================================
        # 👥 DESCRIPTION: Colleague grouping
        # =================================================
        same, prev, nex, g_eve = [], [], [], []

        for nm in names:
            if nm == "Sathiesh M":
                continue

            today_shift = shifts[nm][i]

            if today_shift == sath:
                same.append(nm)

            if today_shift in ("G", "EVE"):
                g_eve.append(f"{nm} ({today_shift})")

        # Rotation logic
        if sath == "S1":
            if i > 0:
                prev = [nm for nm in names if nm != "Sathiesh M" and shifts[nm][i-1] == "S3"]
            nex = [nm for nm in names if nm != "Sathiesh M" and shifts[nm][i] == "S2"]

        elif sath == "S2":
            prev = [nm for nm in names if nm != "Sathiesh M" and shifts[nm][i] == "S1"]
            nex = [nm for nm in names if nm != "Sathiesh M" and shifts[nm][i] == "S3"]

        elif sath == "S3":
            prev = [nm for nm in names if nm != "Sathiesh M" and shifts[nm][i] == "S2"]
            if i < len(dates) - 1:
                nex = [nm for nm in names if nm != "Sathiesh M" and shifts[nm][i+1] == "S1"]

        # Build description text
        desc = ["Colleagues in same shift:"]
        desc += [f"- {n}" for n in same]

        if g_eve:
            desc += ["", "Colleagues in G / EVE:"]
            desc += [f"- {n}" for n in g_eve]

        desc += ["", "Previous shift:"]
        desc += [f"- {n}" for n in prev]

        desc += ["", "Next shift:"]
        desc += [f"- {n}" for n in nex]

        lines.append("DESCRIPTION:" + "\\n".join(desc))

    # =====================================================
    # 🔵 NON-WORKING DAY (OFF / LEAVE)
    # =====================================================
    else:
        start_date = ymd.strftime("%Y%m%d")
        end_date = (ymd + timedelta(days=1)).strftime("%Y%m%d")

        lines.append(f"DTSTART;VALUE=DATE:{start_date}")
        lines.append(f"DTEND;VALUE=DATE:{end_date}")

        lines.append("SUMMARY:" + ("Leave" if sath == "L" else "Off"))

        # List working colleagues
        working = [
            f"{nm} ({shifts[nm][i]})"
            for nm in names
            if shifts[nm][i] in WORKING_SHIFTS
        ]

        desc = ["Working colleagues today:"]
        desc += [f"- {w}" for w in working]

        lines.append("DESCRIPTION:" + "\\n".join(desc))

        # Optional color for OFF day
        if sath == "OFF":
            lines.append("COLOR:#33B679")
            lines.append("X-GOOGLE-CALENDAR-COLOR:#33B679")

    lines.append("END:VEVENT")

# Close calendar
lines.append("END:VCALENDAR")

# =========================================================
# 💾 SAVE FILE
# =========================================================

ics_text = "\n".join(lines)

with open(output_path, "w") as f:
    f.write(ics_text)

print("\n✅ ICS file generated successfully!")
print(f"📁 Saved at: {output_path}")