import pandas as pd

from utils.constants import Constants


class CsvRegister:
    def read_csv_register(self):
        # df_excel = pd.read_excel(Constants.CSV_REG_REPORT)
        df_excel = pd.read_excel(
            str(Constants.CSV_REG_REPORT),  # there are 2 different adp files one for daily and one for bi-weekly
            usecols=[
                Constants.PAYROLL_NAME,
                Constants.UNIQUE_COLUMN,
                'DED CD M'
            ],
            header=0,
            sheet_name=Constants.ADP_SHEET_NAME,
        )
        df_excel = df_excel[df_excel['DED CD M'].notnull()]

        print(df_excel)