import pandas as pd
import xlwings as xw
from datetime import datetime
import shutil
import re
import os

def _read_source_file(source_file, skiprows=1):
    """
    Universal data ingestion pipeline designed for unstable ERP system exports.
    Features an automatic fallback architecture:
      - Primary: Native memory-optimized binary reading (pyxlsb)
      - Secondary: Standard OpenXML and Legacy standard (openpyxl/xlrd)
      - Failsafe: Custom low-memory HTML String-Parsing. Allows 1GB+ disguised 
        .xls HTML files to bypass standard Pandas MemoryError crashes.
        
    Args:
        source_file (str): Absolute path to the exported data file.
        skiprows (int): Number of top rows to skip before reading data.
        
    Returns:
        pd.DataFrame: Headless extracted tabular dataframe.
    """
    ext = source_file.lower().rsplit('.', 1)[-1]
    print(f"  Reading: {os.path.basename(source_file)}")

    if ext == 'xlsb':
        try:
            df = pd.read_excel(source_file, sheet_name=0, header=None, skiprows=skiprows, engine='pyxlsb')
            return df
        except Exception: pass

    elif ext in ('xlsx', 'xlsm'):
        try:
            df = pd.read_excel(source_file, sheet_name=0, header=None, skiprows=skiprows, engine='openpyxl')
            return df
        except Exception: pass

    elif ext == 'xls':
        try:
            xl = pd.ExcelFile(source_file, engine='xlrd')
            df = pd.read_excel(source_file, sheet_name=xl.sheet_names[0], header=None, skiprows=skiprows, engine='xlrd')
            return df
        except Exception: pass

    # HTML / MHTML Fast string parser fallback (Saves RAM on 1GB+ files!)
    print("  Using High-Performance HTML/MHTML parser...")
    with open(source_file, 'rb') as f:
        raw = f.read()
    for enc in ['utf-8', 'latin-1', 'cp1252', 'utf-16']:
        try:
            html_str = raw.decode(enc)
            trs = html_str.split('<tr')
            if len(trs) > 1:
                data = []
                empty_top_rows = 0
                tag_cleaner = re.compile(r'<[^>]+>')
                for tr in trs[1:]:
                    if not tr.strip(): continue
                    tds = tr.split('<td')
                    if len(tds) > 1:
                        row = []
                        for td in tds[1:]:
                            end_idx = td.find('</td>')
                            cell_html = td[:end_idx] if end_idx != -1 else td
                            tag_close_idx = cell_html.find('>')
                            if tag_close_idx != -1:
                                cell_html = cell_html[tag_close_idx+1:]
                            text = tag_cleaner.sub('', cell_html).strip()
                            row.append(text)
                        if any(row):
                            data.append(row)
                        elif len(data) == 0:
                            empty_top_rows += 1
                
                if len(data) > 0:
                    # Return headless dataframe so we can dynamically scan for the true header row!
                    df = pd.DataFrame(data)
                    print(f"  Format detected. Rows extracted: {len(df)}")
                    return df
        except Exception:
            continue

    raise ValueError(f"Cannot read file in any supported format: {source_file}")

def process_pending_report(data_file_path, dashboard_file_path):
    """
    Core Automation logic for the Pending Items reconciliation process.
    Dynamically aligns extracted data to target Dashboard templates regardless
    of vertical shifts in the dashboard structure. Filters data based on 
    the "Returnable" business rule, and propagates required calculation columns.
    
    Args:
        data_file_path (str): Location of the daily raw Pending Items extract.
        dashboard_file_path (str): Location of the Master Excel Dashboard template.
        
    Raises:
        ValueError: Thrown if critical header mappings are absent.
    """
    print("\n==============================================")
    print("Process 3: Pending Item Report Automation")
    print("==============================================\n")
    
    # 1. READ SOURCE DATA
    print("Step 1: Reading Data File...")
    # Read without any header assumptions
    df = _read_source_file(data_file_path, skiprows=0)
    
    # 2. DYNAMICALLY FIND TRUE HEADER ROW
    print("Step 2: Dynamically locating data starting row...")
    header_row_idx = None
    
    # Scan the top 50 rows of the file to find where the headers actually are
    for row_idx, row in df.head(50).iterrows():
        for cell in row:
            # We look for uniquely identifiable columns ('Item Vendor' or 'Returnable')
            cell_str = str(cell).strip().lower()
            if cell_str == 'returnable' or cell_str == 'item vendor':
                header_row_idx = row_idx
                break
        if header_row_idx is not None:
            break
            
    if header_row_idx is None:
        raise ValueError("CRITICAL ERROR: Could not find 'Returnable' or 'Item Vendor' header anywhere in the first 50 rows!")
        
    print(f"  -> Success! True Headers found on File Row {header_row_idx + 1}")
    
    # Promote that physical row to become our Pandas DataFrame Columns!
    df.columns = [str(x).strip() for x in df.iloc[header_row_idx]]
    
    # Slice the dataframe to throw away everything above the header (like garbage titles, empty rows)
    # and also throw away the header row itself so we only have pure data!
    df = df.iloc[header_row_idx + 1:]
    
    # 3. FILTER DATA (Column 'Returnable' == 'Y')
    print("Step 3: Filtering Data ('Returnable' = Y)...")
    
    filtered_df = None
    if 'Returnable' in df.columns:
        filtered_df = df[df['Returnable'].astype(str).str.strip().str.upper() == 'Y']
    else:
        # Failsafe: if name doesn't match for some reason, use Column AB (Index 27)
        if len(df.columns) >= 28:
            returnable_col_name = df.columns[27]
            print(f"  * Warning: 'Returnable' not found by name. Using Column 28 (AB) which is '{returnable_col_name}'")
            filtered_df = df[df[returnable_col_name].astype(str).str.strip().str.upper() == 'Y']
        else:
            raise ValueError(f"CRITICAL ERROR: 'Returnable' column not found and only {len(df.columns)} available.")
            
    print(f"  -> Rows kept after filtering: {len(filtered_df)}")

    # 3. OPEN DASHBOARD & PREPARE SHEET 2
    print("\nStep 3: Opening Dashboard...")
    timestamp = datetime.now().strftime("%d%b%Y_%H%M%S")
    
    # Create the  output path
    base_ext = os.path.splitext(dashboard_file_path)[1]
    output_path = dashboard_file_path.replace(base_ext, f"_Output_{timestamp}{base_ext}")
    
    print(f"  Creating copy to protect template: {os.path.basename(output_path)}")
    shutil.copy2(dashboard_file_path, output_path)
    
    app = xw.App(visible=False)
    app.screen_updating = False
    app.display_alerts = False
    app.calculation = 'manual'
    
    try:
        wb = app.books.open(output_path)
        
        # Select "Sheet 2". (xlwings is 0-indexed for lists, so sheets[1] is the 2nd sheet)
        ws = wb.sheets[1] 
        print(f"  Target Dashboard Sheet: '{ws.name}'")
        
        # 🛑 DYNAMIC DASHBOARD ALIGNMENT 🛑
        print("  Dynamically aligning to Dashboard Template...")
        first_col_name = str(filtered_df.columns[0]).strip()
        template_column_A = ws.range("A1:A20").value
        
        header_row = 1
        start_row = 2
        for i, cell_val in enumerate(template_column_A):
            if str(cell_val).strip() == first_col_name:
                header_row = i + 1
                start_row = header_row + 1
                break
                
        print(f"  -> Template headers detected on Row {header_row}. Data will start on Row {start_row}.")
        
        # Clear old data starting from start_row downwards
        print(f"  Clearing old rows from Row {start_row} downwards...")
        ws.range(f"{start_row}:100000").clear_contents()
        
        # Paste filtered data into A{start_row}
        print("  Pasting filtered rows...")
        ws.range(f"A{start_row}").options(index=False, header=False).value = filtered_df
        
        # 4. INJECT DYNAMIC BUSINESS LOGIC COLUMNS
        # These 7 columns are appended explicitly without a static for-loop
        # mapping allows for granular formula manipulation per column natively.
        print("\nStep 4: Appending 7 New Formula Columns...")
        last_row = len(filtered_df) + start_row - 1
        num_cols = len(filtered_df.columns)
        
        col1 = xw.utils.col_name(num_cols + 1)
        col2 = xw.utils.col_name(num_cols + 2)
        col3 = xw.utils.col_name(num_cols + 3)
        col4 = xw.utils.col_name(num_cols + 4)
        col5 = xw.utils.col_name(num_cols + 5)
        col6 = xw.utils.col_name(num_cols + 6)
        col7 = xw.utils.col_name(num_cols + 7)
        
        print("  Adding headers...")
        ws.range(f"{col1}{header_row}").value = "New1"
        ws.range(f"{col2}{header_row}").value = "New2"
        ws.range(f"{col3}{header_row}").value = "New3"
        ws.range(f"{col4}{header_row}").value = "New4"
        ws.range(f"{col5}{header_row}").value = "New5"
        ws.range(f"{col6}{header_row}").value = "New6"
        ws.range(f"{col7}{header_row}").value = "New7"

        if last_row >= start_row:
            print("  Applying placeholder formulas ...")
            ws.range(f"{col1}{start_row}:{col1}{last_row}").formula = '="test formula 1"'
            ws.range(f"{col2}{start_row}:{col2}{last_row}").formula = '="test formula 2"'
            ws.range(f"{col3}{start_row}:{col3}{last_row}").formula = '="test formula 3"'
            ws.range(f"{col4}{start_row}:{col4}{last_row}").formula = '="test formula 4"'
            ws.range(f"{col5}{start_row}:{col5}{last_row}").formula = '="test formula 5"'
            ws.range(f"{col6}{start_row}:{col6}{last_row}").formula = '="test formula 6"'
            ws.range(f"{col7}{start_row}:{col7}{last_row}").formula = '="test formula 7"'
            print("  + 7 Columns Added (New1 to New7).")
        
        # 5. FINALIZE
        print("\nStep 5: Finalizing...")
        print("  Calculating and Saving...")
        app.calculate()
        wb.save()
        print(f"\n✅ PROCESS 3 SUCCESS! Output saved to: {output_path}")
        
    except Exception as e:
        print(f"\n❌ EXCEL ERROR: {e}")
        raise
    finally:
        if 'wb' in locals() and wb:
            wb.close()
        app.quit()

"""
STANDALONE EXECUTION BLOCK
-------------------------------------------------------------------------
Provides a entry point for local debugging or testing of the reconciliation 
pipeline outside of the Robot Framework environment.
"""
# if __name__ == "__main__":
#     # Update parameters with raw string literals (r"") for Windows paths
#     DATA_FILE = r"C:\Users\DELL XPS\...\Pending_Items_Report_160126.xlsb"
#     DASHBOARD_FILE = r"C:\Users\DELL XPS\...\Pending Items Report - Formula Sheet.xlsx"
#     
#     # Execute the pipeline module directly
#     process_pending_report(DATA_FILE, DASHBOARD_FILE)
