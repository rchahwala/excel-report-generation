from __future__ import print_function, unicode_literals
from collections.abc import Mapping

import datetime

from models.generate_report import GenerateReport
from utils.constants import Constants
from PyInquirer import prompt

# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    DAILY = 'Daily'
    BI_WEEKLY = "Bi-weekly Audit"

    # report selection list
    report_selection = [{
        'type': 'list',
        'name': 'Report',
        'message': 'Please select a type for automatic reconciliation report generation',
        'choices': [DAILY, BI_WEEKLY]
    }]

    selection = prompt(report_selection)

    if selection['Report'] == BI_WEEKLY:
        print("\n")
        print(f"Okay got it! Running {BI_WEEKLY} report. "
              f"Your patient will be test here. Please drink some water and relax"
              )
        print("\n")
        # create the output file with today's date'
        today = datetime.datetime.now().strftime('%m_%d_%Y')
        output_file_name = f"{Constants.OUTPUT_AUDIT_REPORT_PATH}/output_{today}.xlsx"
        print("\n")

        for idx, report in enumerate(Constants.AUDIT_REPORT_TYPES):
            # if idx == 2:
            #     exit(0)

            print(report['oracle_col'])
            reportObj = GenerateReport(
                oracle_col=report['oracle_col'],
                adp_col=report['adp_register_col'],
                output_sheet_name=report['output_sheet_name'],
                output_filename=output_file_name,
                adp_report=Constants.AUDIT_REG_REPORT

            )
            reportObj.generate_report()
            print(f"Sheet generated with sheet name: {report['output_sheet_name']}")
            print("\n")

        print("Audit File Recon Successfully completed")

    if selection['Report'] == DAILY:
        print("\n")
        print(f"Okay got it! Running {DAILY} report. "
              f"Your patient will be test here. Please drink some water and relax"
              )
        print("\n")
        # create the output file with today's date'
        today = datetime.datetime.now().strftime('%m_%d_%Y')
        output_file_name = f"{Constants.OUTPUT_REPORT_PATH}/output_{today}.xlsx"
        print("\n")

        for idx, report in enumerate(Constants.REPORT_TYPES):
            if idx == 2:
                exit(0)

            print(report['oracle_col'])
            reportObj = GenerateReport(
                oracle_col=report['oracle_col'],
                adp_col=report['adp_col'],
                output_sheet_name=report['output_sheet_name'],
                output_filename=output_file_name,
                adp_report=Constants.ADP_REPORT
            )
            reportObj.generate_report()
            print(f"Sheet generated with sheet name: {report['output_sheet_name']}")
            print("\n")

        print("File Recon Successfully completed")
