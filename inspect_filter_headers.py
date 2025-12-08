import pandas as pd

def inspect_filter():
    try:
        df_filter = pd.read_excel('输入数据.xlsx', sheet_name='筛选条件', nrows=5)
        print("Filter Sheet Columns:", df_filter.columns.tolist())
        print("Filter Sheet Data Head:\n", df_filter.head())
    except Exception as e:
        print(e)

if __name__ == "__main__":
    inspect_filter()
