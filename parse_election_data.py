import pandas as pd
import numpy as np
import json
import re
import os

def parse_candidate(cand_str):
    """Extracts candidate name, party, and incumbency status from raw strings."""
    cand_str = str(cand_str).strip()
    
    # Look for party/incumbency inside parentheses e.g. "Candidate Name (D-Inc)"
    match = re.search(r'\((.*?)\)', cand_str)
    if match:
        party_str = match.group(1)
    else:
        # Fallback if no parentheses, takes the last string chunk
        party_str = cand_str.split()[-1]
        
    is_inc = 'INC' in party_str.upper()
    # Clean the extracted string to get just the party character(s)
    party = party_str.upper().replace('-INC', '').replace(' INC', '').replace('INC', '').strip()
    
    # Remove the parentheses and trailing whitespace to get the raw name
    name = re.sub(r'\(.*?\)', '', cand_str).replace('-Inc', '').strip()
    return name, party, is_inc

def main():
    # Defining directory and file paths
    # Use the script's location to build relative paths, making it more portable
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_path = os.path.join(script_dir, "CCTP_data", "CCPT_2024_General_Election.xlsx")
    output_dir = os.path.join(script_dir, "output")

    os.makedirs(output_dir, exist_ok=True)

    xls = pd.ExcelFile(excel_path)
    metadata = {}
    
    for tab in xls.sheet_names:
        if tab in ["About", "Template"]:
            continue
            
        # Read tab without headers to preserve the A column as the index definitions
        df = pd.read_excel(xls, sheet_name=tab, header=None)
        
        # Handle the special "Statewide Results" tab
        if "Statewide" in tab:
            # Start at 2 to skip the empty Timestamp row (index 1)
            end_row = 2
            while end_row < len(df) and not pd.isna(df.iloc[end_row, 0]) and str(df.iloc[end_row, 0]).strip() != "":
                end_row += 1
                
            cols = df.iloc[0:end_row, 0].astype(str).str.strip().tolist()
            cols[0], cols[1] = 'Date', 'Timestamp'
            
            data = df.iloc[0:end_row, 1:].copy()
            data = data.replace(r'(?i)^\s*(unknown)?\s*$', np.nan, regex=True)
            data_T = data.T
            data_T.columns = cols
            data_T['Jurisdiction'] = 'California'
            
            # Format the Date column to only show the date component (YYYY-MM-DD)
            data_T['Date'] = pd.to_datetime(data_T['Date']).dt.date
            
            data_T.to_csv(os.path.join(output_dir, f"{tab}.csv"), index=False)
            
            metadata[tab] = {
                "Race Name": tab,
                "Jurisdictions": ["California"]
            }
            continue

        # Parse Candidate 1 (Row 3, index 2) and Candidate 2 (Row 4, index 3)
        name1, party1, is_inc1 = parse_candidate(df.iloc[2, 0])
        name2, party2, is_inc2 = parse_candidate(df.iloc[3, 0])
        
        # Handle identical parties by appending an increment (D1, D2)
        if party1 == party2:
            party1, party2 = f"{party1}1", f"{party2}2"
            
        incumbent = name1 if is_inc1 else (name2 if is_inc2 else None)
        
        # Find jurisdictions dynamically starting from Row 9 (index 8)
        # The number of header/ID rows before the jurisdictions start
        id_vars = ['Date', 'Timestamp', party1, party2, 'Margin', 'Daily Margin Change', 'Vote Difference', 'Total Votes Cast', 'Total Unprocessed Ballots*']
        jur_start = len(id_vars)
        jur_end = jur_start
        while jur_end < len(df) and not pd.isna(df.iloc[jur_end, 0]) and str(df.iloc[jur_end, 0]).strip() != "":
            jur_end += 1
            
        jurisdictions = df.iloc[jur_start:jur_end, 0].astype(str).str.strip().tolist()
        
        # Update metadata for JSON file
        metadata[tab] = {
            "Race Name": tab,
            "Candidates": [name1, name2],
            "Incumbent": incumbent,
            "Jurisdictions": jurisdictions
        }
        
        # Slice, clean, and transpose the actual values (from Column B onwards)
        data = df.iloc[0:jur_end, 1:].copy()
        data = data.replace(r'(?i)^\s*(unknown)?\s*$', np.nan, regex=True)
        data_T = data.T
        
        # Remap columns and melt Unprocessed Ballots
        data_T.columns = id_vars + jurisdictions
        
        # Format the Date column to only show the date component (YYYY-MM-DD)
        data_T['Date'] = pd.to_datetime(data_T['Date']).dt.date
        
        melted_df = pd.melt(data_T, id_vars=id_vars, value_vars=jurisdictions, var_name='Jurisdiction', value_name='Unprocessed Ballots')
        
        # Save the processed CSV file to the output directory
        melted_df.to_csv(os.path.join(output_dir, f"{tab}.csv"), index=False)
        
    # Save the metadata JSON file to the output directory
    with open(os.path.join(output_dir, "election_metadata.json"), 'w') as f:
        json.dump(metadata, f, indent=4)

    print(f"Successfully parsed {len(xls.sheet_names)} tabs. Files saved in '{output_dir}'.")

if __name__ == "__main__":
    main()