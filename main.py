#%% Risk Management metrics derived from Origami scheduled reports:
    # 1. Number of open workers' compensation claims
    # 2. Number of closed workers' compensation claims
    # 3. Workers' Compensation Claim Closing Ratio
    # 4. Number of adjuster claims
    # 5. Number of adjusters
    # 6. Average Adjuster Caseload

# County Fiscal Calender
# Fiscal year: December 1 through November 30
# Mid-year: December through May
# Year-end: December through November

# Important: The report filename date represents the first day of the following month.
# Therefore, report_date below represents the actual month of the completed data.

# Example:
#   File date = 2026-01-01
#   Data month = December 2025
#   Start date = 2025-12-01
#   End date = 2025-12-31


#%% Imports
import os
from pathlib import Path
import pandas as pd
import roi_common as roi
import re
from dotenv import load_dotenv
load_dotenv()

#%% roi.save_email_attachments function to download attachments from Origami scheduled reports
os.makedirs("email_attachments", exist_ok=True)

sender_name = 'notifications@origamirisk.com'

roi.save_email_attachments(
    sendername_contains=sender_name,
    save_directory="email_attachments"
)


#%% Folder where saved excel attachments are saved
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
excel_files = (
    list(attachments_path.glob("*.xlsx"))
    + list(attachments_path.glob("*.xls"))
)


for file_path in excel_files:

    filename_date = extract_date_from_filename(file_path.name)

    if not filename_date:
        print(f"Could not extract date from filename: {file_path.name}")
        continue

    try:
        # The filename date represents the first day of the following month.
        # Example: 2026-01-01 report = December 2025 data.
        report_date = pd.to_datetime(filename_date) - pd.DateOffset(months=1)

        excel_file = pd.ExcelFile(file_path)

        print(
            f"Processing file: {file_path.name} "
            f"(Data month: {report_date.strftime('%Y-%m')})"
        )

        for sheet_name in excel_file.sheet_names:

            # Read sheet using 5th row as header
            # (skip first 4 rows)
            df = excel_file.parse(
                sheet_name=sheet_name,
                header=4
            )

            # Add the actual month represented by the report
            df["report_date"] = report_date

            # Combine with existing data for the same sheet
            if sheet_name in monthly_data:
                monthly_data[sheet_name] = pd.concat(
                    [monthly_data[sheet_name], df],
                    ignore_index=True
                )
            else:
                monthly_data[sheet_name] = df

            print(
                f"Processed sheet: {sheet_name} "
                f"({len(df)} rows)"
            )

    except Exception as e:
        print(f"Error reading {file_path.name}: {e}")


#%% Create combined data file for each unique sheet name

for sheet_name, df in monthly_data.items():

    print(
        f"Combined data for sheet '{sheet_name}': "
        f"{len(df)} total rows"
    )

    output_file = attachments_path / f"combined_{sheet_name}.xlsx"

    df.to_excel(
        output_file,
        index=False
    )

    print(f"Saved combined data to: {output_file}")


#%% Create dataframes for each new combined file under email_attachments

dataframes = {}


# Function to clean and standardize dataframe keys from filenames
def clean_name(filename: str) -> str:

    name = Path(filename).stem

    # Remove prefix "combined"
    name = re.sub(
        r"^combined[_\- ]*",
        "",
        name,
        flags=re.IGNORECASE
    )

    # Stop before "(ROI)"
    name = name.split("(")[0]

    # Lowercase
    name = name.lower()

    # Replace spaces and hyphens with underscores
    name = re.sub(
        r"[\s\-]+",
        "_",
        name
    )

    # Remove any characters not letters, numbers, or underscore
    name = re.sub(
        r"[^a-z0-9_]",
        "",
        name
    )

    # Remove multiple consecutive underscores
    name = re.sub(
        r"_+",
        "_",
        name
    )

    # Trim leading/trailing underscores
    name = name.strip("_")

    return name


# Load combined files
for file in attachments_path.glob("combined*.xlsx"):

    df_key = clean_name(file.name)

    dataframes[df_key] = pd.read_excel(file)

print("Loaded DataFrames:", list(dataframes.keys()))


#%% Workers' Compensation Metrics

def clean_wc(df):

    df = df.copy()

    # Remove "Grand Totals" rows
    df = df[
        ~df['Coverage'].str.contains(
            "Grand Totals",
            na=False
        )
    ]

    # Ensure report_date is datetime
    df['report_date'] = pd.to_datetime(df['report_date'])

    # Keep only relevant columns
    df = df[
        [
            'report_date',
            'Claim Count'
        ]
    ]

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


# Combine open and closed claims
wc_output = open_monthly.merge(
    closed_monthly,
    on='report_date',
    how='outer'
)


# Calculate monthly closing ratio
wc_output['closing_ratio'] = (
    wc_output['closed_claims'] / wc_output['open_claims']
)


print("\nWorkers' Compensation Monthly Output:")
print(wc_output)


#%% Adjuster Metrics

def clean_adjuster_df(df):

    df = df.copy()

    # Remove "Grand Totals" rows
    df = df[
        ~df['Adjuster User'].str.contains(
            "Grand Totals",
            na=False
        )
    ]

    # Ensure report_date is datetime
    df['report_date'] = pd.to_datetime(df['report_date'])

    # Keep only relevant columns
    df = df[
        [
            'report_date',
            'Adjuster User',
            'Claim Count'
        ]
    ]

    return df


# Apply cleaning function
average_adjuster_caseload = clean_adjuster_df(
    dataframes['average_adjuster_caseload']
)


# Monthly total adjuster claims
monthly_totals = (
    average_adjuster_caseload.groupby('report_date')['Claim Count']
    .sum()
    .reset_index(name='total_claims')
)


# Monthly number of unique adjusters
monthly_adjusters = (
    average_adjuster_caseload.groupby('report_date')['Adjuster User']
    .nunique()
    .reset_index(name='num_adjusters')
)


# Combine adjuster metrics
monthly_caseload = monthly_totals.merge(
    monthly_adjusters,
    on='report_date',
    how='outer'
)


# Monthly average adjuster caseload
monthly_caseload['average_caseload'] = (
    monthly_caseload['total_claims'] / monthly_caseload['num_adjusters']
)


print("\nAdjuster Monthly Output:")
print(monthly_caseload)


#%% Combine monthly metrics

combined_outputs = wc_output.merge(
    monthly_caseload[
        [
            'report_date',
            'total_claims',
            'num_adjusters',
            'average_caseload'
        ]
    ],
    on='report_date',
    how='outer'
)


combined_outputs = combined_outputs.sort_values(
    'report_date'
).reset_index(drop=True)


print("\nCombined Monthly Metrics:")
print(combined_outputs)


#%% Calculate Fiscal Year

def get_fiscal_year(date):

    """
    Fiscal year runs December 1 through November 30.

    Example:
        December 2025 through November 2026 = FY2026
    """

    if date.month == 12:
        return date.year + 1

    return date.year


combined_outputs['fiscal_year'] = (
    combined_outputs['report_date']
    .apply(get_fiscal_year)
)


#%% Calculate Fiscal Period Metrics
def calculate_fiscal_metrics(monthly_df, fiscal_period):

    """
    Calculate fiscal-period metrics only when ALL months required
    for the fiscal period are available in the data.

    Mid-year:
        December through May (6 months required)

    Year-end:
        December through November (12 months required)

    Count metrics are summed.

    Rate/average metrics are recalculated from the
    underlying totals rather than averaging monthly rates.

    If any required month is missing, no fiscal-period
    record is created for that fiscal year.
    """

    results = []

    for fiscal_year, fy_data in monthly_df.groupby('fiscal_year'):

        fy_data = fy_data.sort_values(
            'report_date'
        ).copy()

        # Define the expected months for the fiscal period

        if fiscal_period == 'mid_year':

            # FY2026 = Dec 2025 through May 2026
            expected_dates = pd.date_range(
                start=pd.Timestamp(fiscal_year - 1, 12, 1),
                end=pd.Timestamp(fiscal_year, 5, 1),
                freq='MS'
            )

        elif fiscal_period == 'year_end':

            # FY2026 = Dec 2025 through Nov 2026
            expected_dates = pd.date_range(
                start=pd.Timestamp(fiscal_year - 1, 12, 1),
                end=pd.Timestamp(fiscal_year, 11, 1),
                freq='MS'
            )

        else:
            raise ValueError(
                "fiscal_period must be 'mid_year' or 'year_end'"
            )

        # Normalize report dates to month-start

        fy_data['report_month'] = (
            fy_data['report_date']
            .dt.to_period('M')
            .dt.to_timestamp()
        )

        # Check whether ALL required months are available

        available_dates = set(
            fy_data['report_month']
        )

        missing_dates = [
            date
            for date in expected_dates
            if date not in available_dates
        ]

        if missing_dates:

            print(
                f"Skipping FY{fiscal_year} {fiscal_period}: "
                f"missing month(s): "
                f"{', '.join(d.strftime('%Y-%m') for d in missing_dates)}"
            )

            continue

        # Select only the required months

        period_data = fy_data[
            fy_data['report_month'].isin(expected_dates)
        ].copy()

        # Calculate total count metrics

        total_open_claims = (
            period_data['open_claims'].sum()
        )

        total_closed_claims = (
            period_data['closed_claims'].sum()
        )

        total_adjuster_claims = (
            period_data['total_claims'].sum()
        )

        total_adjusters = (
            period_data['num_adjusters'].sum()
        )

        # Recalculate closing ratio from underlying totals

        if total_open_claims != 0:

            closing_ratio = (
                total_closed_claims /
                total_open_claims
            )

        else:

            closing_ratio = None

        # Recalculate average adjuster caseload

        if total_adjusters != 0:

            average_caseload = (
                total_adjuster_claims /
                total_adjusters
            )

        else:

            average_caseload = None

        # Create fiscal-period record

        results.append({
            'fiscal_year': fiscal_year,
            'period_type': fiscal_period,
            'start_date': expected_dates.min(),
            'end_date': expected_dates.max() + pd.offsets.MonthEnd(0),
            'open_claims': total_open_claims,
            'closed_claims': total_closed_claims,
            'closing_ratio': closing_ratio,
            'total_claims': total_adjuster_claims,
            'average_caseload': average_caseload
        })

    return pd.DataFrame(results)


#%% Calculate Mid-Year Metrics
mid_year_outputs = calculate_fiscal_metrics(
    combined_outputs,
    'mid_year'
)

print("\nFiscal Year Mid-Year Metrics:")
print(mid_year_outputs)


#%% Calculate Year-End Metrics
year_end_outputs = calculate_fiscal_metrics(
    combined_outputs,
    'year_end'
)

print("\nFiscal Year Year-End Metrics:")
print(year_end_outputs)


#%% Convert Fiscal Metrics to Output Format
def format_fiscal_output(fiscal_df):

    if fiscal_df.empty:
        return pd.DataFrame(
            columns=[
                'start_date',
                'end_date',
                'metric',
                'value',
                'fiscal_year',
                'period_type'
            ]
        )

    metric_columns = [
        'open_claims',
        'closed_claims',
        'closing_ratio',
        'total_claims',
        'average_caseload'
    ]

    output = fiscal_df.melt(
        id_vars=[
            'fiscal_year',
            'period_type',
            'start_date',
            'end_date'
        ],
        value_vars=metric_columns,
        var_name='metric',
        value_name='value'
    )

    return output


mid_year_formatted = format_fiscal_output(
    mid_year_outputs
)

year_end_formatted = format_fiscal_output(
    year_end_outputs
)


#%% Add Monthly Output Metadata

combined_outputs['period_type'] = 'monthly'


# Rename report_date to start_date
monthly_formatted = combined_outputs.rename(
    columns={
        'report_date': 'start_date'
    }
).copy()


# Add end_date as end of month
monthly_formatted['end_date'] = (
    monthly_formatted['start_date']
    + pd.offsets.MonthEnd(0)
)


# Keep required output columns
monthly_formatted = monthly_formatted[
    [
        'start_date',
        'end_date',
        'open_claims',
        'closed_claims',
        'closing_ratio',
        'total_claims',
        'num_adjusters',
        'average_caseload',
        'fiscal_year',
        'period_type'
    ]
]


# Melt monthly metrics
monthly_formatted = monthly_formatted.melt(
    id_vars=[
        'start_date',
        'end_date',
        'fiscal_year',
        'period_type'
    ],
    value_vars=[
        'open_claims',
        'closed_claims',
        'closing_ratio',
        'total_claims',
        'num_adjusters',
        'average_caseload'
    ],
    var_name='metric',
    value_name='value'
)


#%% Combine Monthly, Mid-Year, and Year-End Outputs

combined_outputs = pd.concat(
    [
        monthly_formatted,
        mid_year_formatted,
        year_end_formatted
    ],
    ignore_index=True
)


#%% Rename Metrics and Add Metric IDs

metric_name_map = {
    'open_claims': "Number of open workers' compensation claims",
    'closed_claims': "Number of closed workers' compensation claims",
    'closing_ratio': "Workers' Compensation Claim Closing Ratio",
    'total_claims': "Number of adjuster claims",
    'num_adjusters': "Number of adjusters",
    'average_caseload': "Average Adjuster Caseload"
}


metric_id_map = {
    "Workers' Compensation Claim Closing Ratio": 12210,
    "Number of open workers' compensation claims": 102500,
    "Number of closed workers' compensation claims": 102501,
    "Average Adjuster Caseload": 12871,
    "Number of adjuster claims": 102498,
    "Number of adjusters": 102499
}


# Map metric names
combined_outputs['metric'] = (
    combined_outputs['metric'].map(metric_name_map)
)


# Map metric IDs
combined_outputs['metric_id'] = (
    combined_outputs['metric'].map(metric_id_map)
)


#%% Sort Final Output

period_order = {
    'monthly': 1,
    'mid_year': 2,
    'year_end': 3
}

combined_outputs['period_order'] = (
    combined_outputs['period_type'].map(period_order)
)


combined_outputs = combined_outputs.sort_values(
    by=[
        'fiscal_year',
        'end_date',
        'period_order',
        'metric'
    ]
).reset_index(drop=True)

# Drop unneccessary columns for final output merge
combined_outputs = combined_outputs.drop(columns=['period_order', 'fiscal_year', 'period_type'])

# clean the dataframe by dropping rows with any NaN values
final_output = combined_outputs.dropna()

#%% Write to SQL database
roi.roi_merge_metric_data(final_output)