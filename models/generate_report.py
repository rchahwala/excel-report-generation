import os

import pandas as pd

from utils.constants import Constants


class GenerateReport:
    def __init__(
        self, oracle_col, adp_col, output_sheet_name, output_filename, adp_report
    ):
        self.oracle_col = oracle_col
        self.adp_col = adp_col
        self.output_sheet_name = output_sheet_name
        self.output_file_name = output_filename
        self.adp_report = adp_report

    # read oracle report
    def read_oracle_file(self):
        oracle_report = pd.read_excel(
            str(Constants.ORACLE_REPORT),
            usecols=[Constants.UNIQUE_COLUMN, self.oracle_col],
            header=0,
            sheet_name=Constants.ORACLE_SHEET_NAME,
        )

        return oracle_report

    # read adp report
    def read_adp_file(self):
        adp_report = pd.read_excel(
            str(
                self.adp_report
            ),  # there are 2 different adp files one for daily and one for bi-weekly
            usecols=[Constants.PAYROLL_NAME, Constants.UNIQUE_COLUMN, self.adp_col],
            header=0,
            sheet_name=Constants.ADP_SHEET_NAME,
        )
        # line added to handle duplicate values in the csv register report
        adp_report = adp_report[adp_report[self.adp_col].notnull()]
        return adp_report

    # method to check is given file is present in the destination folder
    @staticmethod
    def is_file_present(file_name):
        # check if file is present
        if not os.path.exists(file_name):
            print(
                f"Oh looks like {file_name} is not present in the destination folder. Please do the needful. \n"
            )
            exit("Bye Bye \n")
        else:
            print(f"Great News! {file_name} is present as expected. \n")
            return True

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
            how=Constants.V_LOOKUP_ACTION,
        )

        # create data frame
        df = pd.DataFrame(data=v_lookup)

        # Create index on the unique column
        df = df.set_index(Constants.UNIQUE_COLUMN)

        # replace NaN values with 0 for HCM column
        df[self.oracle_col] = df[self.oracle_col].fillna(0)

        # replace NaN values with 0 for ADP column
        df[self.adp_col] = df[self.adp_col].fillna(0)

        # convert all values to numeric values
        df[[self.oracle_col, self.adp_col]] = df[[self.oracle_col, self.adp_col]].apply(
            pd.to_numeric
        )

        # get delta
        df[Constants.DELTA_COL] = df[self.oracle_col] - df[self.adp_col]

        df[Constants.DELTA_COL].apply(pd.to_numeric)

        # add total and delta
        df.loc["Total"] = pd.Series(
            df.sum(numeric_only=True),
            index=[0, self.oracle_col, self.adp_col, Constants.DELTA_COL],
        )

        with pd.ExcelWriter(
            self.output_file_name, engine="openpyxl", mode="a"
        ) as writer:
            df.to_excel(
                writer,
                sheet_name=self.output_sheet_name,
                na_rep=int(0),
                freeze_panes=(1, 1),
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
