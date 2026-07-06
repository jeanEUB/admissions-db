import pandas as pd
import json
import numpy as np

def convert_value(val):
    if pd.isna(val):
        return ""
    if isinstance(val, (pd.Timestamp, pd.DatetimeTZDtype)):
        # Format as date or datetime
        if val.time() == pd.Timestamp.min.time():
            return val.strftime('%Y-%m-%d')
        else:
            return val.strftime('%Y-%m-%d %H:%M:%S')
    if isinstance(val, float):
        if val.is_integer():
            return int(val)
        return val
    return str(val)

def main():
    excel_path = 'Admissions_Pipeline_Migrated_July_6_v0.3.xlsx'
    
    # 1. Load Dropdown Lists
    xl_file = pd.ExcelFile(excel_path)
    dropdown_df = pd.read_excel(xl_file, sheet_name='Dropdown Lists')
    dropdowns = {}
    for col in dropdown_df.columns:
        options = dropdown_df[col].dropna().astype(str).tolist()
        dropdowns[col] = [opt.strip() for opt in options if opt.strip()]

    # 2. Load Columns and Groups
    # Skip first row which is the sheet title
    df_headers = pd.read_excel(xl_file, sheet_name='Admissions Pipeline', header=None, skiprows=1, nrows=2)
    groups = df_headers.iloc[0].ffill().tolist()
    columns_raw = df_headers.iloc[1].tolist()
    
    # Identify type of each column
    columns_meta = []
    for i, col_name in enumerate(columns_raw):
        col_type = "text"
        is_select = col_name in dropdowns
        
        if is_select:
            col_type = "select"
        elif "date" in col_name.lower() or col_name.lower() in ["created", "modified"]:
            col_type = "date"
        elif "count" in col_name.lower() or col_name.lower() in ["student id", "app serial no."]:
            col_type = "number" if "count" in col_name.lower() else "text"
            
        columns_meta.append({
            "name": col_name,
            "group": groups[i],
            "type": col_type,
            "hasDropdown": is_select
        })
        
    # 3. Load Data Records
    data_df = pd.read_excel(xl_file, sheet_name='Admissions Pipeline', skiprows=2)
    # Ensure column count matches metadata
    data_df.columns = columns_raw
    
    records = []
    for index, row in data_df.iterrows():
        record = {}
        for col_name in columns_raw:
            record[col_name] = convert_value(row[col_name])
        records.append(record)
        
    # 4. Load Dashboard Seed
    dashboard_seed_df = pd.read_excel(xl_file, sheet_name='Dashboard Seed')
    dashboard_info = dashboard_seed_df.dropna(how='all').to_dict(orient='records')

    # Output JS file
    js_content = f"""// Auto-generated data from admissions pipeline Excel spreadsheet
const ADMISSIONS_DATA = {{
    columns: {json.dumps(columns_meta, indent=4)},
    dropdowns: {json.dumps(dropdowns, indent=4)},
    records: {json.dumps(records, indent=4)},
    dashboardSeed: {json.dumps(dashboard_info, indent=4)}
}};

if (typeof module !== 'undefined' && module.exports) {{
    module.exports = ADMISSIONS_DATA;
}}
"""
    
    with open('admissions_data.js', 'w', encoding='utf-8') as f:
        f.write(js_content)
        
    print(f"Data exported successfully! {len(records)} records, {len(columns_meta)} columns, {len(dropdowns)} dropdown lists.")

if __name__ == '__main__':
    main()
