#!/usr/bin/env python3
"""
Check for missing speech therapy sessions within a user-specified date range.

Reads a cleaned Excel file from in/, prompts for school start and end dates,
and produces output files with placeholder rows inserted for every week in
the range that has fewer than 2 sessions.

Usage: python check_sessions.py
"""

import re
import sys
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd
from openpyxl import load_workbook

from parse_speech_logs import (
    style_main_sheet, add_legend_sheet,
)

_SCRIPT_DIR  = Path(__file__).parent
INPUT_FOLDER = _SCRIPT_DIR / 'in'
OUTPUT_FOLDER = _SCRIPT_DIR / 'out'


# ── helpers ───────────────────────────────────────────────────────────────────

def get_date_input(prompt: str) -> datetime:
    """Prompt user for a date in MM/DD/YYYY format."""
    while True:
        raw = input(prompt).strip()
        try:
            return datetime.strptime(raw, '%m/%d/%Y')
        except ValueError:
            print('  Invalid format. Please use MM/DD/YYYY.')


def select_data_file(folder: Path) -> Path:
    """List Excel and CSV files in folder and let user pick one."""
    files = sorted(
        list(folder.glob('*.xlsx')) + list(folder.glob('*.csv')),
        key=lambda f: f.name,
    )
    if not files:
        print(f'Error: no .xlsx or .csv files found in {folder}')
        sys.exit(1)

    print(f'\nData files in {folder.resolve()}:')
    for i, f in enumerate(files, 1):
        print(f'  {i}. {f.name}')

    while True:
        choice = input('\nSelect file number: ').strip()
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(files):
                return files[idx]
        except ValueError:
            pass
        print(f'  Please enter a number between 1 and {len(files)}.')


# ── gap detection for a user-specified range ──────────────────────────────────

def fill_missing_weeks_for_range(
    df: pd.DataFrame, start_date: datetime, end_date: datetime,
) -> pd.DataFrame:
    """Insert placeholder rows for all Mondays in [start, end] with < 2 sessions.

    Unlike the main script's auto-detection (which only looks between the
    first and last session), this covers the full user-specified range so
    missing weeks at the beginning or end are caught.
    """
    start_monday = start_date - timedelta(days=start_date.weekday())
    end_monday   = end_date   - timedelta(days=end_date.weekday())

    all_mondays: list[datetime] = []
    cur = start_monday
    while cur <= end_monday:
        all_mondays.append(cur)
        cur += timedelta(days=7)

    result_rows: list[dict] = []
    for monday in all_mondays:
        ms = monday.strftime('%m/%d/%Y')
        week = df[df['Week of (Monday)'] == ms]

        for _, row in week.iterrows():
            result_rows.append(row.to_dict())

        for _ in range(max(0, 2 - len(week))):
            placeholder = {col: '' for col in df.columns}
            placeholder['Week of (Monday)'] = ms
            result_rows.append(placeholder)

    return pd.DataFrame(result_rows, columns=df.columns) if result_rows else df


# ── save output ───────────────────────────────────────────────────────────────

def save_checked_output(main_df: pd.DataFrame, stem: Path) -> tuple[Path, Path]:
    """Write Excel (with formatting + legend) and CSV. Return (xlsx, csv)."""
    stem.parent.mkdir(parents=True, exist_ok=True)
    xlsx = stem.with_suffix('.xlsx')
    with pd.ExcelWriter(xlsx, engine='openpyxl') as writer:
        main_df.to_excel(writer, index=False, sheet_name='Speech Logs')
        style_main_sheet(writer.sheets['Speech Logs'], main_df)

    wb = load_workbook(xlsx)
    add_legend_sheet(wb)
    wb.save(xlsx)

    csv = stem.with_suffix('.csv')
    main_df.drop(columns=['Goal Category'], errors='ignore').to_csv(csv, index=False)

    return xlsx, csv


# ── entry point ───────────────────────────────────────────────────────────────

def main():
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    data_path = select_data_file(INPUT_FOLDER)
    print(f'\nSelected: {data_path.name}')

    start_date = get_date_input('\nEnter school start date (MM/DD/YYYY): ')
    end_date   = get_date_input('Enter school end date   (MM/DD/YYYY): ')

    if end_date < start_date:
        print('Error: end date must be after start date.')
        sys.exit(1)

    if data_path.suffix.lower() == '.csv':
        df = pd.read_csv(data_path)
    else:
        df = pd.read_excel(data_path, sheet_name='Speech Logs')
    print(f'\nLoaded {len(df)} rows from {data_path.name}')

    # process each child independently
    children = [c for c in df['Child Name'].dropna().unique() if c]
    if not children:
        print('Error: no child names found in the spreadsheet.')
        sys.exit(1)

    parts = [
        fill_missing_weeks_for_range(
            df[df['Child Name'] == child].copy(), start_date, end_date,
        )
        for child in children
    ]
    result = pd.concat(parts, ignore_index=True)

    added = len(result) - len(df)

    safe_name = re.sub(r'[^\w\s-]', '', data_path.stem).strip().replace(' ', '_')
    today = datetime.today().strftime('%Y-%m-%d')
    stem = OUTPUT_FOLDER / f'{safe_name}_checked_{today}'

    xlsx, csv = save_checked_output(result, stem)

    print(f'\n  Added {added} placeholder row(s) for missing sessions')
    print(f'  Output Excel : {xlsx.name}')
    print(f'  Output CSV   : {csv.name}')
    print()


if __name__ == '__main__':
    main()
