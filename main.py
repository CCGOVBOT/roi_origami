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

            # Add a column with the report date
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

# Dictionary to store dataframes
dataframes = {}

# Function to clean and standardize dataframe keys from filenames
def clean_name(filename: str) -> str:
    # Remove extension
    name = Path(filename).stem

    # Remove prefix "combined"
    name = re.sub(r"^combined[_\- ]*", "", name, flags=re.IGNORECASE)

    # Stop before "(ROI)"
    name = name.split("(")[0]

    # Lowercase
    name = name.lower()

    # Replace spaces, hyphens with underscores
    name = re.sub(r"[\s\-]+", "_", name)

    # Remove any characters not letters, numbers, or underscore
    name = re.sub(r"[^a-z0-9_]", "", name)

    # Remove multiple consecutive underscores
    name = re.sub(r"_+", "_", name)

    # Trim leading/trailing underscores
    name = name.strip("_")

    return name

# Load combined files
for file in attachments_path.glob("combined*.xlsx"):
    df_key = clean_name(file.name)
    dataframes[df_key] = pd.read_excel(file)

print("Loaded DataFrames:", list(dataframes.keys()))

# %% Metrics in scope [with relevant dataframe(s)]:
    #1   Workers' Compensation Claim Closing Ratio ['closed_wc', 'open_wc']
    #2   Worker's Compensation Closed ['closed_wc']
    #3   Worker's Compensation New Opens ['open_wc']
    #4   Average Adjuster Caseload ['average_adjuster_caseload']

# Metrics #1-3: Workers' Compensation 
# Function to clean workers' compensation dataframes
def clean_wc(df):
    df = df.copy()
    
    # Remove "Grand Total" rows
    df = df[~df['Coverage'].str.contains("Grand Totals", na=False)]
    
    # Ensure date is datetime
    df['report_date'] = pd.to_datetime(df['report_date'])
    
    # Keep only relevant columns
    df = df[['report_date', 'Claim Count']]
    
    return df

# Apply clean_wc function to both dataframes
open_wc = clean_wc(dataframes['open_wc'])
closed_wc = clean_wc(dataframes['closed_wc'])

# Monthly totals for open claims
open_monthly = (
    open_wc.groupby('report_date')['Claim Count']
    .sum()
    .reset_index(name='open_claims')
)

# Monthly totals for closed claims
closed_monthly = (
    closed_wc.groupby('report_date')['Claim Count']
    .sum()
    .reset_index(name='closed_claims')
)

wc_output = open_monthly.merge(closed_monthly, on='report_date', how='outer')

# Calculate closing ratio = closed claims / open claims
wc_output['closing_ratio'] = wc_output['closed_claims'] / wc_output['open_claims']

print(wc_output)


# %%
# Metric #4: Average Adjuster Caseload ['average_adjuster_caseload']

# Function to clean average adjuster caseload dataframe
def clean_adjuster_df(df):
    df = df.copy()
    
    # Remove "Grand Totals" rows
    df = df[~df['Adjuster User'].str.contains("Grand Totals", na=False)]
    
    # Ensure report_date is datetime
    df['report_date'] = pd.to_datetime(df['report_date'])
    
    # Keep only relevant columns
    df = df[['report_date', 'Adjuster User', 'Claim Count']]
    
    return df

# Apply cleaning function
average_adjuster_caseload = clean_adjuster_df(dataframes['average_adjuster_caseload'])

# Calculate monthly totals
monthly_totals = (
    average_adjuster_caseload.groupby('report_date')['Claim Count']
    .sum()
    .reset_index(name='total_claims')
)

# Calculate number of unique adjusters per month
monthly_adjusters = (
    average_adjuster_caseload.groupby('report_date')['Adjuster User']
    .nunique()
    .reset_index(name='num_adjusters')
)

# Merge claim and adjuster count to calculate average caseload
monthly_caseload = monthly_totals.merge(monthly_adjusters, on='report_date')
monthly_caseload['average_caseload'] = monthly_caseload['total_claims'] / monthly_caseload['num_adjusters']

print(monthly_caseload)

# %% create combined_outputs merging wc_output and monthly_caseload

# merge dataframes on report_date
combined_outputs = wc_output.merge(
    monthly_caseload[['report_date', 'average_caseload']],
    on='report_date',
    how='left'  # keep all months from final
)

# transpose data so that each metric is a row
combined_outputs = combined_outputs.melt(id_vars=['report_date'], var_name='metric', value_name='value')

# change report_date to start_date
combined_outputs.rename(columns={'report_date': 'start_date'}, inplace=True)

# add in end_date as end of month
combined_outputs['end_date'] = combined_outputs['start_date'] + pd.offsets.MonthEnd(0)

print(combined_outputs)

# %% rename metrics in metric column to full metric name and add in metric ID column

metric_name_map = {
    'open_claims': "Worker's Compensation New Opens",
    'closed_claims': "Worker's Compensation Closed",
    'closing_ratio': "Workers' Compensation Claim Closing Ratio",
    'average_caseload': "Average Adjuster Caseload"
}

metric_id_map = {
    "Workers' Compensation Claim Closing Ratio": 12210,
    "Worker's Compensation Closed": 12496,
    "Worker's Compensation New Opens": 12495,
    "Average Adjuster Caseload": 12871
}

combined_outputs['metric'] = combined_outputs['metric'].map(metric_name_map)
combined_outputs['metric_id'] = combined_outputs['metric'].map(metric_id_map)

print(combined_outputs)


# %% Write to sql database
roi.roi_merge_metric_data(combined_outputs)
# %%
