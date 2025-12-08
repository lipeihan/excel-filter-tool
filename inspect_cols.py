import pandas as pd

def inspect_cols():
    df_hours_1 = pd.read_excel('输入数据.xlsx', sheet_name='累计工时', nrows=0)
    print("Sheet '累计工时' columns:", df_hours_1.columns.tolist())
    
    df_hours_2 = pd.read_excel('输入数据.xlsx', sheet_name='工时数据', nrows=0)
    print("Sheet '工时数据' columns:", df_hours_2.columns.tolist())

if __name__ == "__main__":
    inspect_cols()
