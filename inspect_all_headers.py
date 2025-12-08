import pandas as pd

def inspect_all():
    xl = pd.ExcelFile('输入数据.xlsx')
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, nrows=0)
        print(f"--- Sheet: {sheet} ---")
        print(df.columns.tolist())
        print()

if __name__ == "__main__":
    inspect_all()
