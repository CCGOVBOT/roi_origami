#%% Imports
import os
from pathlib import Path
import pandas as pd
import roi_common as roi
import re
from dotenv import load_dotenv
load_dotenv()

#%% roi.save_email_attachments function
os.makedirs("email_attachments", exist_ok=True)

sender_name = 'notifications@origamirisk.com'

roi.save_email_attachments(sendername_contains=sender_name, save_directory="email_attachments")

#%% Folder where ssaved excel attachments are saved
attachments_path = Path("email_attachments")

# Create dictionary to store combined sheet data by sheet name
monthly_data = {}  # { sheet_name: dataframe }

# Function to extract date from filename (YYYY-MM-DD)
def extract_date_from_filename(filename):
    match = re.search(r'\d{4}-\d{2}-\d{2}', filename)
    if match:
        return match.group(0)
    return None

# Get all Excel files
excel_files = list(attachments_path.glob("*.xlsx")) + list(attachments_path.glob("*.xls"))

for file_path in excel_files:
    report_date = extract_date_from_filename(file_path.name)
    if not report_date:
        print(f"Could not extract date from filename: {file_path.name}")
        continue

    try:
        excel_file = pd.ExcelFile(file_path)
        print(f"Processing file: {file_path.name} (Report date: {report_date})")

        for sheet_name in excel_file.sheet_names:
            # Read sheet, using 5th row as header (skip first 4 rows)
            df = excel_file.parse(sheet_name=sheet_name, header=4)

            # Optionally add a column with the report date
            df["report_date"] = report_date

            # Combine with existing data for the same sheet
            if sheet_name in monthly_data:
                monthly_data[sheet_name] = pd.concat(
                    [monthly_data[sheet_name], df],
                    ignore_index=True
                )
            else:
                monthly_data[sheet_name] = df

            print(f"Processed sheet: {sheet_name} ({len(df)} rows)")

    except Exception as e:
        print(f"Error reading {file_path.name}: {e}")

#%% 
# Create combined data file for each unique sheet name to consolidate metric logic as new monthly data is added
for sheet_name, df in monthly_data.items():
    print(f"Combined data for sheet '{sheet_name}': {len(df)} total rows")
    # save combined data to a new Excel file
    output_file = attachments_path / f"combined_{sheet_name}.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Saved combined data to: {output_file}")


# %% Create dataframes for each new combined file under email_attachments
