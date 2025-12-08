import pandas as pd

file = '输出格式.xlsx'

print(f"--- Inspecting {file} ---")
try:
    df = pd.read_excel(file)
    print(f"Columns: {list(df.columns)}")
    print("First 5 rows:")
    print(df.head().to_string())
    
    # 尝试读取更多行以防筛选条件写在其他地方，比如第二行或者特定的说明区域
    # 如果有多个sheet，也检查一下
    xl = pd.ExcelFile(file)
    print(f"\nSheet names: {xl.sheet_names}")
    if len(xl.sheet_names) > 1:
        for sheet in xl.sheet_names[1:]:
             print(f"\n--- Inspecting Sheet: {sheet} ---")
             df_sheet = pd.read_excel(file, sheet_name=sheet)
             print(f"Columns: {list(df_sheet.columns)}")
             print(df_sheet.head().to_string())

except Exception as e:
    print(f"Error reading {file}: {e}")
