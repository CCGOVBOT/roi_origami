#%% Imports
import os
from pathlib import Path
import pandas as pd
import roi_common as roi

from dotenv import load_dotenv
load_dotenv()
load_dotenv('/mnt/repos/ROI/.env')

#%% roi.save_email_attachments function
os.makedirs("email_attachments", exist_ok=True)

sender_name = 'origamirisk'

roi.save_email_attachments(sendername_contains=sender_name, save_directory="email_attachments")

#%% Read saved attachments - to be updated
# Create dictionary to store dataframes for each email match string
email_dataframes = {}

# List of email match strings to process
email_match_strings = {
    'fund_activity_match_string': fund_activity_match_string,
    'tender_totals_match_string': tender_totals_match_string
}
# Process each email match string
for key, match_string in email_match_strings.items():
    matching_files = []
    
    # Find all CSV files in email_attachments directory that contain the match string
    for file in Path("email_attachments").glob("*.csv"):
        if match_string in file.name:
            matching_files.append(file)
    
    # Read and concatenate all matching files
    if matching_files:
        df_list = []
        for file_path in matching_files:
            try:
                # Try different encodings in order
                for encoding in ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding, skiprows=3)
                        df_list.append(df)
                        print(f"Loaded {file_path.name} for {key} (encoding: {encoding})")
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    print(f"Error reading {file_path.name}: Could not decode with any standard encoding")
            except Exception as e:
                print(f"Error reading {file_path.name}: {e}")
        
        # Concatenate and deduplicate
        if df_list:
            combined_df = pd.concat(df_list, ignore_index=True)
            combined_df = combined_df.drop_duplicates()
            email_dataframes[key] = combined_df
            print(f"Combined {len(matching_files)} file(s) into dataframe for {key} ({len(combined_df)} rows)")
    else:
        print(f"No files found matching '{match_string}'")

# Assign to individual dataframes for easier access
fund_activity = email_dataframes.get('fund_activity_match_string')
tender_totals = email_dataframes.get('tender_totals_match_string')
