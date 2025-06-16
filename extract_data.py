import pandas as pd
import numpy as np
import os
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

"""
Portuguese Tourism Data Extractor - Updated Version
Handles all years 2014-2023 with specialized extraction logic:
- 2014-2015: Q1, Q2... sheet format
- 2016: 611, 612... sheet format
- 2017-2021, 2023: Standard 2.1 sheet format
- 2022: Sheet 3.1 format (tourism data moved to section 3)
"""

def extract_2014_2015_format(xlsx_file, year):
    """
    Extract data from 2014-2015 format using Q1, Q2, Q3... sheets
    """
    sheet_names = xlsx_file.sheet_names

    # Look for Q1 sheet (should contain main tourism indicators)
    target_sheet = None
    for name in ['Q1', 'Quadro 1', 'q1']:
        if name in sheet_names:
            target_sheet = name
            break

    if not target_sheet:
        return None

    df = pd.read_excel(xlsx_file, sheet_name=target_sheet, header=None)

    # Search for tourism guest data
    for idx, row in df.iterrows():
        if idx >= 50:  # Don't search too far
            break

        first_cell = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""

        # Look for patterns indicating total guests
        if any(pattern in first_cell for pattern in ['total', 'hóspedes', 'hospedes']):
            # Check if this has monthly data (12+ numeric values)
            numeric_count = 0
            monthly_values = []

            for col_idx in range(1, min(20, len(row))):
                try:
                    val = row.iloc[col_idx]
                    if pd.notna(val) and isinstance(val, (int, float)) and val > 0:
                        numeric_count += 1
                        monthly_values.append(val)
                except:
                    pass

            if numeric_count >= 6:  # At least 6 months of data
                return {
                    'row_idx': idx,
                    'pattern': first_cell,
                    'monthly_values': monthly_values[:12],  # Take first 12 values
                    'sheet': target_sheet
                }

    return None

def extract_2016_format(xlsx_file, year):
    """
    Extract data from 2016 format using 611, 612... sheets
    """
    sheet_names = xlsx_file.sheet_names

    # Look for 611 sheet (should be equivalent to 2.1)
    target_sheet = None
    for name in ['611', '6.1.1', '61.1']:
        if name in sheet_names:
            target_sheet = name
            break

    if not target_sheet:
        return None

    df = pd.read_excel(xlsx_file, sheet_name=target_sheet, header=None)

    # Search for tourism guest data (similar logic to 2014-2015)
    for idx, row in df.iterrows():
        if idx >= 50:
            break

        first_cell = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""

        if any(pattern in first_cell for pattern in ['total', 'hóspedes', 'hospedes']):
            numeric_count = 0
            monthly_values = []

            for col_idx in range(1, min(20, len(row))):
                try:
                    val = row.iloc[col_idx]
                    if pd.notna(val) and isinstance(val, (int, float)) and val > 0:
                        numeric_count += 1
                        monthly_values.append(val)
                except:
                    pass

            if numeric_count >= 6:
                return {
                    'row_idx': idx,
                    'pattern': first_cell,
                    'monthly_values': monthly_values[:12],
                    'sheet': target_sheet
                }

    return None

def extract_2022_format(xlsx_file, year):
    """
    Extract data from 2022 - tourism data is in sheet 3.1, row 9
    Based on analysis report: sheet 3.1, row 9 contains "Total" with tourism values
    """
    sheet_names = xlsx_file.sheet_names

    # For 2022, tourism data moved to sheet 3.1
    if '3.1' not in sheet_names:
        print(f"   ❌ Sheet 3.1 not found in 2022 file")
        return None

    try:
        df = pd.read_excel(xlsx_file, sheet_name='3.1', header=None)

        # Based on analysis, row 9 contains "Total" with tourism data
        # Structure: [Total, annual_total, jan, feb, mar, ..., dec]
        target_row = 9

        if target_row >= len(df):
            print(f"   ❌ Row {target_row} not found in sheet 3.1")
            return None

        row = df.iloc[target_row]
        first_cell = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""

        # Verify this is the correct row (should contain "total")
        if 'total' not in first_cell:
            print(f"   ❌ Row {target_row} doesn't contain 'total': '{first_cell}'")
            # Try to find the correct row nearby
            for search_row in range(max(0, target_row-3), min(len(df), target_row+5)):
                search_cell = str(df.iloc[search_row, 0]).strip().lower() if pd.notna(df.iloc[search_row, 0]) else ""
                if 'total' in search_cell and len(search_cell) < 20:  # Avoid long descriptions
                    target_row = search_row
                    row = df.iloc[target_row]
                    first_cell = search_cell
                    print(f"   ✅ Found 'total' at row {target_row}")
                    break
            else:
                return None

        # Extract monthly data from the row
        # Expected format: [Label, Annual_Total, Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec]
        monthly_values = []
        annual_total = None

        # Try to get annual total from column 1
        try:
            annual_val = row.iloc[1]
            if pd.notna(annual_val) and isinstance(annual_val, (int, float)) and annual_val > 1000:
                annual_total = annual_val
        except:
            pass

        # Extract 12 monthly values starting from column 2
        for month_idx in range(12):
            col_idx = 2 + month_idx  # Start from column 2
            if col_idx < len(row):
                try:
                    val = row.iloc[col_idx]
                    if pd.notna(val) and isinstance(val, (int, float)):
                        monthly_values.append(val)
                    else:
                        monthly_values.append(np.nan)
                except:
                    monthly_values.append(np.nan)
            else:
                monthly_values.append(np.nan)

        # Validate the data looks like tourism data (not business data)
        valid_values = [v for v in monthly_values if pd.notna(v) and v > 0]
        if valid_values:
            avg_val = sum(valid_values) / len(valid_values)
            # Tourism data should be in hundreds to thousands range, not single digits
            if avg_val < 50:  # Business ratios/percentages are usually small
                print(f"   ❌ Values look like business data, not tourism (avg: {avg_val:.1f})")
                return None

        # Calculate annual total if not found
        if annual_total is None and valid_values:
            annual_total = sum(valid_values)

        print(f"   ✅ Extracted tourism data from sheet 3.1, row {target_row}")
        print(f"       Annual total: {annual_total}")
        print(f"       Sample values: {valid_values[:6]}")

        return {
            'row_idx': target_row,
            'pattern': first_cell,
            'monthly_values': monthly_values,
            'annual_total': annual_total,
            'sheet': '3.1'
        }

    except Exception as e:
        print(f"   ❌ Error reading sheet 3.1: {e}")
        return None

def extract_standard_format(xlsx_file, year):
    """
    Extract data from standard format (2017-2021, 2023) using 2.1 sheet
    """
    if '2.1' not in xlsx_file.sheet_names:
        return None

    df = pd.read_excel(xlsx_file, sheet_name='2.1', header=None)

    # Look for "Total" row around row 9-11
    for idx in range(8, min(15, len(df))):
        row = df.iloc[idx]
        first_cell = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""

        if first_cell.lower() == 'total':
            # Extract monthly data starting from column 2
            monthly_values = []
            annual_total = None

            # Annual total usually in column 1
            try:
                annual_total = row.iloc[1]
                if not isinstance(annual_total, (int, float)):
                    annual_total = None
            except:
                annual_total = None

            # Monthly data starting from column 2
            for col_idx in range(2, min(14, len(row))):
                try:
                    val = row.iloc[col_idx]
                    if pd.notna(val) and isinstance(val, (int, float)):
                        monthly_values.append(val)
                    else:
                        monthly_values.append(np.nan)
                except:
                    monthly_values.append(np.nan)

            # Ensure we have 12 values
            while len(monthly_values) < 12:
                monthly_values.append(np.nan)

            return {
                'row_idx': idx,
                'pattern': 'total',
                'monthly_values': monthly_values[:12],
                'annual_total': annual_total,
                'sheet': '2.1'
            }

    return None

def extract_tourism_data_specialized(sources_folder='sources', output_folder='data'):
    """
    Specialized extraction that handles all different formats
    """
    Path(output_folder).mkdir(exist_ok=True)

    all_data = []
    extraction_log = []

    years = range(2014, 2024)

    print("SPECIALIZED TOURISM DATA EXTRACTION")
    print("=" * 50)

    for year in years:
        # Determine file extension and name
        if year <= 2015:
            filename = f"ET_{year}.xls"
        else:
            filename = f"ET_{year}.xlsx"

        filepath = Path(sources_folder) / filename

        if not filepath.exists():
            print(f"⚠️  Warning: {filename} not found, skipping...")
            extraction_log.append(f"{year}: FILE NOT FOUND")
            continue

        print(f"📊 Processing {filename}...")

        try:
            xlsx_file = pd.ExcelFile(filepath)
            result = None

            # Use year-specific extraction logic
            if year in [2014, 2015]:
                print(f"   🔍 Using 2014-2015 format (Q1, Q2... sheets)")
                result = extract_2014_2015_format(xlsx_file, year)
            elif year == 2016:
                print(f"   🔍 Using 2016 format (611, 612... sheets)")
                result = extract_2016_format(xlsx_file, year)
            elif year == 2022:
                print(f"   🔍 Using 2022 special format (sheet 3.1)")
                result = extract_2022_format(xlsx_file, year)
                if result:
                    print(f"   ✅ Found 2022 data in sheet {result['sheet']}")
                else:
                    print(f"   ❌ Failed to find 2022 data in sheet 3.1")
            else:
                print(f"   🔍 Using standard format (2.1 sheet)")
                result = extract_standard_format(xlsx_file, year)

            if not result:
                print(f"❌ Could not extract data for {year}")
                extraction_log.append(f"{year}: NO DATA FOUND")
                continue

            print(f"   ✅ Found data in sheet '{result['sheet']}', row {result['row_idx']}")
            print(f"       Pattern: '{result['pattern']}'")

            # Create year data entry
            annual_total = result.get('annual_total')
            if annual_total is None:
                # Calculate from monthly data
                valid_monthly = [x for x in result['monthly_values'] if not pd.isna(x)]
                annual_total = sum(valid_monthly) if valid_monthly else None

            year_data = {
                'Year': year,
                'Annual_Total': annual_total,
                'Extraction_Method': result['pattern'],
                'Source_Row': result['row_idx'],
                'Source_Sheet': result['sheet']
            }

            # Add monthly data
            months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                     'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

            for month, value in zip(months, result['monthly_values']):
                year_data[month] = value if not pd.isna(value) else None

            all_data.append(year_data)

            # Count valid monthly values
            valid_months = sum(1 for v in result['monthly_values'] if not pd.isna(v) and v > 0)
            print(f"   📈 Extracted {valid_months}/12 monthly values")
            extraction_log.append(f"{year}: SUCCESS - {valid_months}/12 months from {result['sheet']}")

        except Exception as e:
            print(f"❌ Error processing {filename}: {str(e)}")
            extraction_log.append(f"{year}: ERROR - {str(e)}")
            continue

    if not all_data:
        print("❌ No data was extracted successfully!")
        return None, extraction_log

    # Convert to DataFrame and process
    df_annual = pd.DataFrame(all_data)

    # Create monthly time series format
    monthly_data = []
    for _, row in df_annual.iterrows():
        year = int(row['Year'])
        for month_idx, month in enumerate(['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                                          'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'], 1):
            if pd.notna(row[month]) and row[month] is not None:
                monthly_data.append({
                    'Date': f"{year}-{month_idx:02d}-01",
                    'Year': year,
                    'Month': month_idx,
                    'Month_Name': month,
                    'Guests_Thousands': row[month]
                })

    df_monthly = pd.DataFrame(monthly_data)
    if not df_monthly.empty:
        df_monthly['Date'] = pd.to_datetime(df_monthly['Date'])
        df_monthly = df_monthly.sort_values('Date').reset_index(drop=True)

    # Save files
    print("\n" + "=" * 50)
    print("💾 Saving specialized extraction results...")

    # Save annual summary
    annual_file = Path(output_folder) / 'portugal_tourism_annual_specialized.csv'
    df_annual.to_csv(annual_file, index=False)
    print(f"✅ Annual data: {annual_file}")

    # Save monthly time series
    if not df_monthly.empty:
        monthly_file = Path(output_folder) / 'portugal_tourism_monthly_specialized.csv'
        df_monthly.to_csv(monthly_file, index=False)
        print(f"✅ Monthly data: {monthly_file}")

        # Create R-ready file
        ts_df = df_monthly[['Date', 'Guests_Thousands']].copy()
        ts_df.columns = ['Date', 'Guests']
        ts_file = Path(output_folder) / 'portugal_tourism_ts_specialized.csv'
        ts_df.to_csv(ts_file, index=False)
        print(f"✅ R-ready file: {ts_file}")

    # Save extraction log
    log_df = pd.DataFrame({'Year': range(2014, 2024), 'Status': extraction_log})
    log_file = Path(output_folder) / 'specialized_extraction_log.csv'
    log_df.to_csv(log_file, index=False)
    print(f"✅ Extraction log: {log_file}")

    # Print final summary
    print("\n" + "=" * 50)
    print("📈 SPECIALIZED EXTRACTION SUMMARY")
    print("=" * 50)

    if not df_monthly.empty:
        print(f"Total observations: {len(df_monthly)}")
        print(f"Years covered: {', '.join(map(str, sorted(df_monthly['Year'].unique())))}")
        print(f"Date range: {df_monthly['Date'].min().strftime('%Y-%m')} to {df_monthly['Date'].max().strftime('%Y-%m')}")
        print(f"Missing values: {df_monthly['Guests_Thousands'].isna().sum()}")

        if len(df_monthly) > 0:
            print(f"Mean monthly guests: {df_monthly['Guests_Thousands'].mean():.1f} thousand")
            peak_idx = df_monthly['Guests_Thousands'].idxmax()
            low_idx = df_monthly['Guests_Thousands'].idxmin()
            print(f"Peak: {df_monthly.loc[peak_idx, 'Date'].strftime('%Y-%m')} ({df_monthly.loc[peak_idx, 'Guests_Thousands']:.1f}k)")
            print(f"Low: {df_monthly.loc[low_idx, 'Date'].strftime('%Y-%m')} ({df_monthly.loc[low_idx, 'Guests_Thousands']:.1f}k)")

    print("\nExtraction details:")
    for year_log in extraction_log:
        if 'SUCCESS' in year_log:
            print(f"  ✅ {year_log}")
        elif 'ERROR' in year_log:
            print(f"  ❌ {year_log}")
        else:
            print(f"  ⚠️  {year_log}")

    return df_monthly, df_annual, extraction_log

if __name__ == "__main__":
    # Run specialized extraction
    monthly_df, annual_df, log = extract_tourism_data_specialized()

    if monthly_df is not None and not monthly_df.empty:
        print("\n🎉 SPECIALIZED EXTRACTION COMPLETED!")
        print("Files ready for R analysis:")
        print("  - portugal_tourism_ts_specialized.csv")
        print("\nNext: Load this file in R for time series analysis")
    else:
        print("❌ Specialized extraction failed")
