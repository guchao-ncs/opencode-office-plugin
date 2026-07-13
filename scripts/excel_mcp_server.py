# /// script
# dependencies = [
#   "fastmcp",
#   "pandas",
#   "openpyxl",
#   "tabulate",
# ]
# ///

import sys
import os
import platform
import json
import subprocess
from fastmcp import FastMCP

# Create the MCP server instance
mcp = FastMCP("excel")

def get_filename_from_stack():
    # Helper to check if any context variables point to a file, if needed
    return None

def normalize_2d(raw, row_count, col_count):
    if raw is None:
        return [["" for _ in range(col_count)] for _ in range(row_count)]
    # If it's a single value (not a list)
    if not isinstance(raw, list):
        # Repeat it to fit, or return it wrapped
        return [[raw]]
    # If it's a 1D list (either 1 row or 1 col)
    if len(raw) > 0 and not isinstance(raw[0], list):
        if row_count == 1:
            return [raw]
        elif col_count == 1:
            return [[val] for val in raw]
        else:
            return [raw]
    return raw

def run_jxa_code(js_code, args=[]):
    cmd = ["osascript", "-l", "JavaScript", "-e", js_code] + [str(arg) for arg in args]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        raise Exception(f"JXA error: {res.stderr.strip()}")
    return res.stdout.strip()

# Windows COM client imports and functions
def read_active_sheet_win():
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return {"error": "Microsoft Excel is not running or active"}
    
    try:
        wb = excel.ActiveWorkbook
        if not wb:
            return {"error": "No active workbook"}
        sheet = excel.ActiveSheet
        if not sheet:
            return {"error": "No active sheet"}
        
        sheet_name = sheet.Name
        try:
            used_range = sheet.UsedRange
        except Exception:
            return {"sheetName": sheet_name, "empty": True}
            
        if not used_range:
            return {"sheetName": sheet_name, "empty": True}
            
        address = used_range.Address
        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count
        start_row = used_range.Row
        start_col = used_range.Column
        
        values = used_range.Value
        formulas = used_range.Formula
        
        # win32com tuples need conversion to lists for JSON serialization
        if isinstance(values, tuple):
            values = [list(row) for row in values]
        if isinstance(formulas, tuple):
            formulas = [list(row) for row in formulas]
            
        return {
            "sheetName": sheet_name,
            "address": address,
            "rowCount": row_count,
            "colCount": col_count,
            "startRow": start_row,
            "startCol": start_col,
            "values": values,
            "formulas": formulas
        }
    finally:
        pythoncom.CoUninitialize()

def write_cell_win(cell_address: str, value: str):
    import win32com.client
    import pythoncom
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return {"error": "Microsoft Excel is not running or active"}
        
    try:
        wb = excel.ActiveWorkbook
        if not wb:
            return {"error": "No active workbook"}
        sheet = excel.ActiveSheet
        if not sheet:
            return {"error": "No active sheet"}
            
        cell = sheet.Range(cell_address)
        if value.startswith("="):
            cell.Formula = value
        else:
            try:
                if '.' in value:
                    cell.Value = float(value)
                else:
                    cell.Value = int(value)
            except ValueError:
                cell.Value = value
        return {"success": True, "address": cell_address, "value": value}
    finally:
        pythoncom.CoUninitialize()


@mcp.tool()
def read_active_sheet() -> str:
    """Reads the used range data, formulas, and cells from the active worksheet in the running Excel instance."""
    system = platform.system()
    if system == "Darwin":
        js_code = """
        function run() {
            try {
                const excel = Application("Microsoft Excel");
                if (!excel.running()) {
                    return JSON.stringify({ error: "Microsoft Excel is not running" });
                }
                if (excel.workbooks.length === 0) {
                    return JSON.stringify({ error: "No workbooks open" });
                }
                const activeSheet = excel.activeWorkbook.activeSheet;
                const sheetName = activeSheet.name();
                let usedRange;
                try {
                    usedRange = activeSheet.usedRange;
                } catch (e) {
                    return JSON.stringify({ sheetName: sheetName, empty: true });
                }
                if (!usedRange) {
                    return JSON.stringify({ sheetName: sheetName, empty: true });
                }
                const address = usedRange.address();
                const rowCount = usedRange.rows.count();
                const colCount = usedRange.columns.count();
                const startRow = usedRange.row();
                const startCol = usedRange.column();
                const values = usedRange.value();
                const formulas = usedRange.formula();
                return JSON.stringify({
                    sheetName: sheetName,
                    address: address,
                    rowCount: rowCount,
                    colCount: colCount,
                    startRow: startRow,
                    startCol: startCol,
                    values: values,
                    formulas: formulas
                });
            } catch (e) {
                return JSON.stringify({ error: e.message });
            }
        }
        """
        try:
            output = run_jxa_code(js_code)
            data = json.loads(output)
            if "error" in data:
                return json.dumps(data, indent=2)
            if data.get("empty"):
                return f"Worksheet '{data['sheetName']}' is empty."
            
            rc = data["rowCount"]
            cc = data["colCount"]
            data["values"] = normalize_2d(data.get("values"), rc, cc)
            data["formulas"] = normalize_2d(data.get("formulas"), rc, cc)
            return json.dumps(data, indent=2)
        except Exception as e:
            return json.dumps({"error": f"Failed to run JXA on macOS: {str(e)}"}, indent=2)
            
    elif system == "Windows":
        try:
            data = read_active_sheet_win()
            if "error" in data:
                return json.dumps(data, indent=2)
            if data.get("empty"):
                return f"Worksheet '{data['sheetName']}' is empty."
            
            rc = data["rowCount"]
            cc = data["colCount"]
            data["values"] = normalize_2d(data.get("values"), rc, cc)
            data["formulas"] = normalize_2d(data.get("formulas"), rc, cc)
            return json.dumps(data, indent=2)
        except Exception as e:
            return json.dumps({"error": f"Failed to run win32com on Windows: {str(e)}"}, indent=2)
    else:
        return json.dumps({"error": f"Unsupported platform: {system}"}, indent=2)


@mcp.tool()
def write_cell(cell_address: str, value: str) -> str:
    """Writes a formula or value to a specific cell (e.g. 'C4') in the active worksheet of the running Excel instance."""
    system = platform.system()
    if system == "Darwin":
        js_code = """
        function run(argv) {
            try {
                const cellAddr = argv[0];
                const cellValue = argv[1];
                const excel = Application("Microsoft Excel");
                if (!excel.running()) {
                    return JSON.stringify({ error: "Microsoft Excel is not running" });
                }
                if (excel.workbooks.length === 0) {
                    return JSON.stringify({ error: "No workbooks open" });
                }
                const activeSheet = excel.activeWorkbook.activeSheet;
                const cell = activeSheet.range(cellAddr);
                if (cellValue.startsWith("=")) {
                    cell.formula = cellValue;
                } else {
                    if (!isNaN(cellValue) && cellValue.trim() !== "") {
                        cell.value = Number(cellValue);
                    } else {
                        cell.value = cellValue;
                    }
                }
                return JSON.stringify({ success: true, address: cellAddr, value: cellValue });
            } catch (e) {
                return JSON.stringify({ error: e.message });
            }
        }
        """
        try:
            output = run_jxa_code(js_code, [cell_address, value])
            return output
        except Exception as e:
            return json.dumps({"error": f"Failed to run JXA on macOS: {str(e)}"}, indent=2)
            
    elif system == "Windows":
        try:
            result = write_cell_win(cell_address, value)
            return json.dumps(result, indent=2)
        except Exception as e:
            return json.dumps({"error": f"Failed to run win32com on Windows: {str(e)}"}, indent=2)
    else:
        return json.dumps({"error": f"Unsupported platform: {system}"}, indent=2)


@mcp.tool()
def analyze_excel_file(file_path: str) -> str:
    """Inspects a local Excel or CSV file and returns sheet names and column shapes/types using pandas."""
    import pandas as pd
    
    if not os.path.exists(file_path):
        return f"Error: File not found at {file_path}"
        
    _, ext = os.path.splitext(file_path.lower())
    
    try:
        if ext == '.csv':
            df = pd.read_csv(file_path)
            info = {
                "file_type": "CSV",
                "shape": df.shape,
                "columns": [
                    {"name": col, "type": str(dtype)} 
                    for col, dtype in zip(df.columns, df.dtypes)
                ]
            }
            return json.dumps(info, indent=2)
        elif ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
            xl = pd.ExcelFile(file_path)
            sheets_info = {}
            for sheet_name in xl.sheet_names:
                df = xl.parse(sheet_name)
                sheets_info[sheet_name] = {
                    "shape": df.shape,
                    "columns": [
                        {"name": col, "type": str(dtype)} 
                        for col, dtype in zip(df.columns, df.dtypes)
                    ]
                }
            info = {
                "file_type": "Excel",
                "sheets": xl.sheet_names,
                "sheets_details": sheets_info
            }
            return json.dumps(info, indent=2)
        else:
            return f"Unsupported file extension: {ext}. Only CSV and Excel (.xlsx, .xls, .xlsm, .xlsb) files are supported."
    except Exception as e:
        return f"Error analyzing file: {str(e)}"


@mcp.tool()
def run_pandas_query(file_path: str, query: str, sheet_name: str = None) -> str:
    """Runs a pandas DataFrame query expression or aggregation (e.g. `df.query('col > 10')` or `df.groupby('col').mean()`) on a local Excel/CSV file and returns the result. Use 'df' as the DataFrame variable name in your query."""
    import pandas as pd
    
    if not os.path.exists(file_path):
        return f"Error: File not found at {file_path}"
        
    _, ext = os.path.splitext(file_path.lower())
    
    try:
        if ext == '.csv':
            df = pd.read_csv(file_path)
        elif ext in ['.xlsx', '.xls', '.xlsm', '.xlsb']:
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
        else:
            return f"Unsupported file type: {ext}"
            
        query_str = query.strip()
        local_vars = {"df": df, "pd": pd}
        result = eval(query_str, globals(), local_vars)
        
        if isinstance(result, pd.DataFrame):
            try:
                return result.to_markdown(index=False)
            except Exception:
                return result.to_string(index=False)
        elif isinstance(result, pd.Series):
            try:
                return result.to_markdown()
            except Exception:
                return result.to_string()
        else:
            return str(result)
            
    except Exception as e:
        return f"Error executing query: {str(e)}"

if __name__ == "__main__":
    mcp.run()
