import os
import glob
import pyexcel
from tkinter import Tk, filedialog

# === 1️⃣ Setup interactive folder and file name selection ===

print("📂 Please select the folder with your .xls files...")
Tk().withdraw()  # Hide the main Tkinter window
source_folder = filedialog.askdirectory(title="Select folder with .xls files")

if not source_folder:
    print("❌ No folder selected. Exiting.")
    exit(1)

# Ask user for file name
output_name = input("💾 Enter a name for the combined output file (without extension): ").strip()

# Ensure it ends with .xls
if not output_name.lower().endswith(".xls"):
    output_name += ".xls"

# Set output path inside a subfolder called 'output'
output_dir = os.path.join(source_folder, 'output')
output_file = os.path.join(output_dir, output_name)

# Create output directory if needed
os.makedirs(output_dir, exist_ok=True)

# === 2️⃣ Load and combine all .xls files ===

xls_files = glob.glob(os.path.join(source_folder, '*.xls'))

if not xls_files:
    print(f"❌ No .xls files found in {source_folder}")
    exit(1)

combined_data = []

# Adjust as needed
expected_columns = ['* login', '* organisation', '* roles', '* add to lessons [yes/no]']
expected_norm = [col.lower().strip() for col in expected_columns]

for file in xls_files:
    print(f'📄 Reading: {os.path.basename(file)}')

    sheet = pyexcel.get_sheet(file_name=file)
    num_rows = sheet.number_of_rows()
    print(f"   ↳ Rows: {num_rows}")

    headers = [h.lower().strip() for h in sheet.row[0]]
    if headers != expected_norm:
        print(f"⚠️  WARNING: Header mismatch in file: {file}")

    data_rows = sheet.row[1:]  # Skip header
    combined_data.extend(data_rows)

# === 3️⃣ Save the combined file ===

# Insert the expected headers at the top
combined_data.insert(0, expected_columns)

pyexcel.save_as(array=combined_data, dest_file_name=output_file, sheet_name="Sheet1")
print(f'✅ Combined .xls created at:\n{output_file}')
