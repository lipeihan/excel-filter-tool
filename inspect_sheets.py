import pandas as pd

def inspect_sheets():
    xl = pd.ExcelFile('输入数据.xlsx')
    print("Sheet Names:", xl.sheet_names)

if __name__ == "__main__":
    inspect_sheets()
