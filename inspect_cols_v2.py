import pandas as pd

def inspect_cols():
    try:
        df_roster = pd.read_excel('输入数据.xlsx', sheet_name='花名册', nrows=0)
        print("Sheet '花名册' columns:", df_roster.columns.tolist())
    except Exception as e:
        print(f"Error reading '花名册': {e}")

    try:
        df_basic = pd.read_excel('输入数据.xlsx', sheet_name='基本数据', nrows=0)
        print("Sheet '基本数据' columns:", df_basic.columns.tolist())
    except Exception as e:
        print(f"Error reading '基本数据': {e}")

if __name__ == "__main__":
    inspect_cols()
