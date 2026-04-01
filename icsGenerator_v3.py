import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime, timedelta
import pandas as pd  # ✅ FIX 1: Added missing import

# =========================================================
# 🖥️ UI SETUP - File selection dialogs
# =========================================================

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

# --- Select destination folder ---
save_folder = filedialog.askdirectory(
    title="Select Folder to Save ICS File"
)

if not save_folder:
    print("No folder selected. Exiting...")
    exit()

# =========================================================
# 📁 FILE NAMING
# =========================================================

current_month = datetime.now().strftime("%b%y")
filename = f"ShiftRooster_{current_month}.ics"
output_path = os.path.join(save_folder, filename)

# =========================================================
# ⚙️ HELPER FUNCTIONS & SHIFT CONFIG
# =========================================================

def dtfmt(dt):
    return dt.strftime("%Y%m%dT%H%M%S")

shift_time = {
    "S1": ("06:00", "15:30"),
    "G": ("09:00", "18:30"),
    "S2": ("14:00", "23:30"),
    "EVE": ("17:00", "02:30"),
    "S3": ("22:00", "07:30")
}

WORKING_SHIFTS = {"S1", "G", "S2", "EVE", "S3"}

# =========================================================
# 📊 FIXED RANGE PARSER (STABLE & SIMPLE)
# =========================================================

df = pd.read_excel(excel_path, header=None)

# ✅ Define structure manually
DATE_ROW = 0
DATA_START_ROW = 3
NAME_COL = 0
START_COL = 0

# Extract dates
raw_dates = df.iloc[DATE_ROW, START_COL:].tolist()

# Extract data block
data = df.iloc[DATA_START_ROW:].copy()

# Drop empty rows
data = data[data.iloc[:, NAME_COL].notna()]

# Extract names
names = data.iloc[:, NAME_COL].astype(str).str.strip().tolist()

print("✅ Detected names:", names)

# Convert dates
dates = []
for d in raw_dates:
    try:
        formatted = pd.to_datetime(d).strftime("%d-%b")
    except:
        formatted = str(d).strip()
    dates.append(formatted)

# Build shifts dictionary
shifts = {}

for i, name in enumerate(names):
    shifts[name] = []

    for j in range(len(dates)):
        value = data.iloc[i, START_COL + j]

        clean_value = str(value).strip().upper()

        if clean_value in ["NAN", ""]:
            clean_value = "OFF"

        shifts[name].append(clean_value)


# =========================================================
# 🔍 VALIDATION
# =========================================================

TARGET_NAME = "Sathiesh M"

if TARGET_NAME not in shifts:
    print(f"❌ Name '{TARGET_NAME}' not found in Excel.")
    print("Available names:", list(shifts.keys()))
    exit()

# =========================================================
# 📅 ICS GENERATION
# =========================================================

lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//Shift Calendar//Sathiesh M//EN"
]

year = datetime.now().year  # ✅ FIX 4: dynamic year

for i, d in enumerate(dates):

    sath = shifts[TARGET_NAME][i]
    ymd = datetime.strptime(f"{d}-{year}", "%d-%b-%Y")

    lines.append("BEGIN:VEVENT")

    # ================= WORKING DAY =================
    if sath in shift_time:

        st, et = shift_time[sath]

        start_dt = datetime.strptime(f"{ymd:%Y-%m-%d} {st}", "%Y-%m-%d %H:%M")
        end_dt = datetime.strptime(f"{ymd:%Y-%m-%d} {et}", "%Y-%m-%d %H:%M")

        if sath in ("S3", "EVE"):
            end_dt += timedelta(days=1)

        lines.append(f"DTSTART;TZID=Asia/Kolkata:{dtfmt(start_dt)}")
        lines.append(f"DTEND;TZID=Asia/Kolkata:{dtfmt(end_dt)}")
        lines.append(f"SUMMARY:{sath} ({st}–{et})")

        # --- Description ---
        same, prev, nex, g_eve = [], [], [], []

        for nm in names:
            if nm == TARGET_NAME:
                continue

            today_shift = shifts[nm][i]

            if today_shift == sath:
                same.append(nm)

            if today_shift in ("G", "EVE"):
                g_eve.append(f"{nm} ({today_shift})")

        # Rotation logic
        if sath == "S1":
            if i > 0:
                prev = [nm for nm in names if nm != TARGET_NAME and shifts[nm][i-1] == "S3"]
            nex = [nm for nm in names if nm != TARGET_NAME and shifts[nm][i] == "S2"]

        elif sath == "S2":
            prev = [nm for nm in names if nm != TARGET_NAME and shifts[nm][i] == "S1"]
            nex = [nm for nm in names if nm != TARGET_NAME and shifts[nm][i] == "S3"]

        elif sath == "S3":
            prev = [nm for nm in names if nm != TARGET_NAME and shifts[nm][i] == "S2"]
            if i < len(dates) - 1:
                nex = [nm for nm in names if nm != TARGET_NAME and shifts[nm][i+1] == "S1"]

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

    # ================= OFF / LEAVE =================
    else:
        start_date = ymd.strftime("%Y%m%d")
        end_date = (ymd + timedelta(days=1)).strftime("%Y%m%d")

        lines.append(f"DTSTART;VALUE=DATE:{start_date}")
        lines.append(f"DTEND;VALUE=DATE:{end_date}")

        lines.append("SUMMARY:" + ("Leave" if sath == "L" else "Off"))

        working = [
            f"{nm} ({shifts[nm][i]})"
            for nm in names
            if shifts[nm][i] in WORKING_SHIFTS
        ]

        desc = ["Working colleagues today:"]
        desc += [f"- {w}" for w in working]

        lines.append("DESCRIPTION:" + "\\n".join(desc))

        if sath == "OFF":
            lines.append("COLOR:#33B679")
            lines.append("X-GOOGLE-CALENDAR-COLOR:#33B679")

    lines.append("END:VEVENT")

lines.append("END:VCALENDAR")

# =========================================================
# 💾 SAVE FILE
# =========================================================

ics_text = "\n".join(lines)

with open(output_path, "w") as f:
    f.write(ics_text)

print("\n✅ ICS file generated successfully!")
print(f"📁 Saved at: {output_path}")