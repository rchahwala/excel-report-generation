# This is a sample Python script.
import datetime

from models.generate_report import GenerateReport
from utils.constants import Constants

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    # create the output file with today's date'
    today = datetime.datetime.now().strftime('%m_%d_%Y')
    output_file_name = f"{Constants.OUTPUT_REPORT_PATH}/output_{today}.xlsx"
    # print(output_file_name)
    print("\n")

    for idx, report in enumerate(Constants.REPORT_TYPES):
        # if idx == 2:
        #     exit(0)

        print(report['oracle_col'])
        reportObj = GenerateReport(
            oracle_col=report['oracle_col'],
            adp_col=report['adp_col'],
            output_sheet_name=report['output_sheet_name'],
            output_filename=output_file_name
        )
        reportObj.generate_report()
        print(f"Sheet generated with sheet name: {report['output_sheet_name']}")
        print("\n")

    print("File Recon Successfully completed")
