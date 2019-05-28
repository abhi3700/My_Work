import xlwings as xw
import pandas as pd


def main():
    # wb = xw.Book.caller()
    # wb.sheets[0].range("A1").value = "Hello xlwings!"     # test code

    # sht_1 = wb.sheets[0]
    # sht_run = wb.sheets['RUN']

    excel_file_directory = 'data/sample.xlsx'
    excel_file = pd.ExcelFile(excel_file_directory)
    df = excel_file.parse('Table 2', skiprows=0)                            # copy a sheet and paste into another sheet and skiprows 8
    # sht_run.range('A1').options(index=False).value = df
    # print(df.columns[0])
    # print(df.loc[[df.columns[0] == 'MENU ITEM']])
    df.columns = ['MENU ITEM', 'Celery', 'Cereals with Gluten', 'Crustaceans', 'Egg', 'Fish', 'Lupin', 'Milk', 'Molluscs', 
                'Mustard', 'Nuts', 'Peanuts', 'Sesame', 'Soybeans', 'Sulphur Dioxide/ Sulphites', 'Garlic', 'Onions', 'Vegans', 'Vegetarians', '-']
    print(df)

if __name__ == '__main__':
    main()