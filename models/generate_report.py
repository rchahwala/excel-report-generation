import os

import numpy as np
import pandas as pd

from utils.constants import Constants


class GenerateReport:

    def __init__(self, oracle_col, adp_col, output_sheet_name, output_filename, adp_report):
        self.oracle_col = oracle_col
        self.adp_col = adp_col
        self.output_sheet_name = output_sheet_name
        self.output_file_name = output_filename
        self.adp_report = adp_report

    # read oracle report
    def read_oracle_file(self):
        oracle_report = pd.read_excel(
            str(Constants.ORACLE_REPORT),
            usecols=[
                Constants.UNIQUE_COLUMN,
                self.oracle_col
            ],
            header=0,
            sheet_name=Constants.ORACLE_SHEET_NAME,
        )

        return oracle_report

    # read adp report
    def read_adp_file(self):
        adp_report = pd.read_excel(
            str(self.adp_report),  # there are 2 different adp files one for daily and one for bi-weekly
            usecols=[
                Constants.PAYROLL_NAME,
                Constants.UNIQUE_COLUMN,
                self.adp_col
            ],
            header=0,
            sheet_name=Constants.ADP_SHEET_NAME,
        )
        return adp_report

    # generate report
    def generate_report(self):
        # write to the output file
        if not os.path.exists(self.output_file_name):
            df = pd.DataFrame()
            df.to_excel(self.output_file_name)
            print("Output File created : ", self.output_file_name)

        # perform v lookup
        v_lookup = pd.merge(
            self.read_adp_file(),
            self.read_oracle_file(),
            on=Constants.UNIQUE_COLUMN,
            how=Constants.V_LOOKUP_ACTION
        )

        # create data frame
        df = pd.DataFrame(data=v_lookup)

        # get delta
        df[Constants.DELTA_COL] = df[self.oracle_col] - df[self.adp_col]

        # add total and delta
        df.loc['Total'] = pd.Series(
            df.sum(),
            index=[
                None,
                self.oracle_col,
                self.adp_col,
                Constants.DELTA_COL
            ]
        )

        with pd.ExcelWriter(self.output_file_name, engine='openpyxl', mode='a') as writer:
            df.to_excel(
                writer,
                sheet_name=self.output_sheet_name,
                na_rep=int(0)
            )

        # with pd.ExcelWriter(self.output_file_name, engine='xlsxwriter') as writer:
        #     # Convert the dataframe to an XlsxWriter Excel object.
        #     df.to_excel(writer, sheet_name='Sheet1', index=False)
        #     # Get the xlsxwriter objects from the dataframe writer object.
        #     # writer.book()
        #     worksheet = writer.sheets['Sheet1']
        #     # Apply the auto filter based on the dimensions of the dataframe.
        #     worksheet.autofilter(0, 0, df.shape[0], df.shape[1] - 1)
        #     # workbook.close()
