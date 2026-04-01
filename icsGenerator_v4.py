import pandas as pd

# =========================================================
# 📊 LOAD CORRECT SHEET (IMPORTANT FIX)
# =========================================================

excel_path = "your_file.xlsx"   # keep your existing path

# ✅ Read ONLY the April'26 sheet (NOT Sheet1)
df = pd.read_excel(excel_path, sheet_name="April'26", header=None)

# =========================================================
# 📊 FIXED RANGE PARSER (STABLE)
# =========================================================

DATE_ROW = 0          # Row where dates are present
DATA_START_ROW = 3    # Row where employee data starts (adjust if needed)
NAME_COL = 0
START_COL = 1

# Extract dates
raw_dates = df.iloc[DATE_ROW, START_COL:].tolist()

# Extract data block
data = df.iloc[DATA_START_ROW:].copy()

# Remove empty name rows
data = data[data.iloc[:, NAME_COL].notna()]

# Extract names
names = data.iloc[:, NAME_COL].astype(str).str.strip().tolist()

print("✅ Detected names:", names)

# =========================================================
# 📅 DATE FORMATTING
# =========================================================

dates = []
for d in raw_dates:
    try:
        formatted = pd.to_datetime(d).strftime("%d-%b")
    except:
        formatted = str(d).strip()
    dates.append(formatted)

# =========================================================
# 🔁 BUILD SHIFT DATA
# =========================================================

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
# 🎯 TARGET NAME CHECK
# =========================================================

TARGET_NAME = "Sathiesh M"   # ⚠️ Change if needed

if TARGET_NAME not in shifts:
    print(f"❌ Name '{TARGET_NAME}' not found in Excel.")
    print("Available names:", list(shifts.keys()))
    exit()

print(f"✅ Generating calendar for {TARGET_NAME}")