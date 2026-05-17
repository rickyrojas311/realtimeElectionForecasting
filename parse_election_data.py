import pandas as pd
import numpy as np
import json
import re
import os
import argparse
import datetime as _dt


def parse_candidate(cand_str):
    """Extracts candidate name, party, and incumbency status from raw strings."""
    cand_str = str(cand_str).strip()
    match = re.search(r'\((.*?)\)', cand_str)
    if match:
        party_str = match.group(1)
    else:
        party_str = cand_str.split()[-1]
    is_inc = 'INC' in party_str.upper()
    party = party_str.upper().replace('-INC', '').replace(' INC', '').replace('INC', '').strip()
    name = re.sub(r'\(.*?\)', '', cand_str).replace('-Inc', '').strip()
    return name, party, is_inc


def _is_2022_format(df):
    """True when row 1, col 0 holds a candidate name (2022 layout)."""
    if len(df) < 2:
        return False
    cell = str(df.iloc[1, 0]).strip()
    return bool(re.search(r'\([DR]', cell, re.IGNORECASE))



def _parse_2022_tab(df, tab, output_dir, metadata, all_races_data):
    """Process a single tab in 2022 format."""
    name1, party1, is_inc1 = parse_candidate(df.iloc[1, 0])
    name2, party2, is_inc2 = parse_candidate(df.iloc[2, 0])

    if party1 == party2:
        party1, party2 = f"{party1}1", f"{party2}2"

    incumbent = name1 if is_inc1 else (name2 if is_inc2 else None)

    # Find jur_start: the row after the 'Unprocessed Ballots' stat row
    jur_start = None
    for i in range(3, len(df)):
        if 'Unprocessed Ballots' in str(df.iloc[i, 0]):
            jur_start = i + 1
            break
    if jur_start is None:
        print(f"  Skipping '{tab}': could not locate 'Unprocessed Ballots' row.")
        return

    # Build id_vars from actual row labels (use party symbols for candidate rows)
    stat_labels = [str(df.iloc[i, 0]).strip() for i in range(3, jur_start)]
    id_vars = ['Date', party1, party2] + stat_labels

    # Jurisdictions: rows jur_start .. first NaN/empty
    jur_end = jur_start
    while jur_end < len(df) and not pd.isna(df.iloc[jur_end, 0]) and str(df.iloc[jur_end, 0]).strip() != '':
        jur_end += 1

    jurisdictions = df.iloc[jur_start:jur_end, 0].astype(str).str.strip().tolist()

    metadata[tab] = {
        "Race Name": tab,
        "Candidates": [name1, name2],
        "Incumbent": incumbent,
        "Jurisdictions": jurisdictions,
    }

    # Raw date values from row 0 (cols 1 onwards), trimming trailing NaN
    raw_dates = df.iloc[0, 1:].tolist()
    last_valid = len(raw_dates) - 1
    while last_valid >= 0 and pd.isna(raw_dates[last_valid]):
        last_valid -= 1
    raw_dates = raw_dates[: last_valid + 1]
    n_cols = len(raw_dates)

    # Slice, clean, and transpose
    data = df.iloc[0:jur_end, 1: n_cols + 1].copy()
    data = data.replace(r'(?i)^\s*(unknown)?\s*$', np.nan, regex=True)
    data_T = data.T
    data_T.columns = id_vars + jurisdictions

    # Ensure 'Total Votes Cast' exists; compute as R + D when the tab omits it
    if 'Total Votes Cast' not in id_vars:
        insert_at = len(id_vars) - 1  # just before the Unprocessed Ballots stat
        id_vars.insert(insert_at, 'Total Votes Cast')
        data_T.insert(
            insert_at,
            'Total Votes Cast',
            pd.to_numeric(data_T[party1], errors='coerce') + pd.to_numeric(data_T[party2], errors='coerce'),
        )

    # Resolve all dates to proper datetimes so string-only drops get a YYYY-MM-DD date too.
    # Infer the year from the first datetime object in the row.
    year = next(
        (v.year for v in raw_dates if isinstance(v, (_dt.datetime, pd.Timestamp))),
        _dt.datetime.now().year,
    )
    parsed_dates: list[_dt.date] = []
    for v in raw_dates:
        if isinstance(v, (_dt.datetime, pd.Timestamp)):
            parsed_dates.append(v.date())
        elif isinstance(v, str):
            date_part = v.split('\n')[0].strip()
            try:
                parsed_dates.append(
                    pd.to_datetime(f"{date_part}-{year}", format='%d-%b-%Y').date()
                )
            except Exception:
                parsed_dates.append(None)  # type: ignore[arg-type]
        else:
            parsed_dates.append(None)  # type: ignore[arg-type]

    data_T['Date'] = [d.strftime('%Y-%m-%d') if d else '' for d in parsed_dates]

    # Dedup: when a day appears twice (morning + afternoon drop), keep only the first
    seen: set = set()
    keep = []
    for i, d in enumerate(parsed_dates):
        if d not in seen:
            seen.add(d)
            keep.append(i)
    data_T = data_T.iloc[keep].reset_index(drop=True)

    melted_df = pd.melt(
        data_T,
        id_vars=id_vars,
        value_vars=jurisdictions,
        var_name='Jurisdiction',
        value_name='Unprocessed Ballots',
    )

    melted_df.to_csv(os.path.join(output_dir, f"{tab}.csv"), index=False)

    if tab.startswith(("CD", "SD", "AD")) and "D" in (party1, party2) and "R" in (party1, party2):
        race_df = melted_df.copy()
        race_df['Race'] = tab
        all_races_data.append(race_df)


def main():
    parser = argparse.ArgumentParser(description="Parse CCTP election Excel data into CSVs.")
    parser.add_argument(
        '--input', '-i',
        help="Path to input Excel file (default: CCTP_data/CCPT_2024_General_Election.xlsx)",
    )
    parser.add_argument(
        '--output', '-o',
        help="Output directory (default: auto-derived from the year in the filename)",
    )
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))

    excel_path = args.input if args.input else os.path.join(
        script_dir, "CCTP_data", "CCPT_2024_General_Election.xlsx"
    )

    if args.output:
        output_dir = args.output
    else:
        basename = os.path.splitext(os.path.basename(excel_path))[0]
        year_match = re.search(r'20\d\d', basename)
        year = year_match.group(0) if year_match else 'election'
        output_dir = os.path.join(script_dir, f"{year}_output")

    os.makedirs(output_dir, exist_ok=True)

    xls = pd.ExcelFile(excel_path)
    metadata = {}
    all_races_data = []

    for tab in xls.sheet_names:
        tab = str(tab)
        if tab in ["About", "Template"]:
            continue

        df = pd.read_excel(xls, sheet_name=tab, header=None)

        # --- 2022 format ---
        if _is_2022_format(df):
            _parse_2022_tab(df, tab, output_dir, metadata, all_races_data)
            continue

        # --- 2024 format: Statewide tab ---
        if "Statewide" in tab:
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
            data_T['Date'] = pd.to_datetime(data_T['Date']).dt.date

            data_T.to_csv(os.path.join(output_dir, f"{tab}.csv"), index=False)
            metadata[tab] = {"Race Name": tab, "Jurisdictions": ["California"]}
            continue

        # --- 2024 format: regular race tab ---
        name1, party1, is_inc1 = parse_candidate(df.iloc[2, 0])
        name2, party2, is_inc2 = parse_candidate(df.iloc[3, 0])

        if party1 == party2:
            party1, party2 = f"{party1}1", f"{party2}2"

        incumbent = name1 if is_inc1 else (name2 if is_inc2 else None)

        id_vars = ['Date', 'Timestamp', party1, party2, 'Margin', 'Daily Margin Change', 'Vote Difference', 'Total Votes Cast', 'Total Unprocessed Ballots*']
        jur_start = len(id_vars)
        jur_end = jur_start
        while jur_end < len(df) and not pd.isna(df.iloc[jur_end, 0]) and str(df.iloc[jur_end, 0]).strip() != "":
            jur_end += 1

        jurisdictions = df.iloc[jur_start:jur_end, 0].astype(str).str.strip().tolist()

        metadata[tab] = {
            "Race Name": tab,
            "Candidates": [name1, name2],
            "Incumbent": incumbent,
            "Jurisdictions": jurisdictions,
        }

        data = df.iloc[0:jur_end, 1:].copy()
        data = data.replace(r'(?i)^\s*(unknown)?\s*$', np.nan, regex=True)
        data_T = data.T
        data_T.columns = id_vars + jurisdictions
        data_T['Date'] = pd.to_datetime(data_T['Date']).dt.date

        melted_df = pd.melt(data_T, id_vars=id_vars, value_vars=jurisdictions, var_name='Jurisdiction', value_name='Unprocessed Ballots')
        melted_df.to_csv(os.path.join(output_dir, f"{tab}.csv"), index=False)

        if tab.startswith(("CD", "SD", "AD")) and "D" in (party1, party2) and "R" in (party1, party2):
            race_df = melted_df.copy()
            race_df['Race'] = tab
            all_races_data.append(race_df)

    with open(os.path.join(output_dir, "election_metadata.json"), 'w') as f:
        json.dump(metadata, f, indent=4)

    print(f"Successfully parsed {len(xls.sheet_names)} tabs. Files saved in '{output_dir}'.")

    if all_races_data:
        combined_df = pd.concat(all_races_data, ignore_index=True)
        combined_df.to_csv(os.path.join(output_dir, "combined_races.csv"), index=False)


if __name__ == "__main__":
    main()
