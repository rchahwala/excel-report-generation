from __future__ import print_function, unicode_literals
from collections.abc import Mapping

import datetime

from models.generate_report import GenerateReport
from utils.constants import Constants
from PyInquirer import prompt

# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    DAILY = 'Daily Report'
    BI_WEEKLY = "Bi-weekly Audit Report"

    # Handle report type selection by the user
    report_selection = [{
        'type': 'list',
        'name': 'Report',
        'message': 'Please select a report type to be generated!',
        'choices': [DAILY, BI_WEEKLY]
    }]

    # selected value
    selection = prompt(report_selection)

    # Instruct users for the next steps
    print("Thanks for your selection! \n")
    print(f"You have selected {selection['Report']} to be generated. \n")

    # handle bi-weekly report generation
    if selection['Report'] == BI_WEEKLY:

        # check if HCM Excel is present in the destination folder
        hcm_excel_is_present = GenerateReport.is_file_present(
            Constants.ORACLE_REPORT
        )

        # check if HCM Excel is present in the destination folder
        csv_register_excel_is_present = GenerateReport.is_file_present(
            Constants.CSV_REG_REPORT
        )

        if hcm_excel_is_present and csv_register_excel_is_present:

            print(f"Next, your patient will be tested now. Please drink some water and relax \n")

            # create the output file with today's date'
            today = datetime.datetime.now().strftime('%m_%d_%Y')
            output_file_name = f"{Constants.OUTPUT_REPORT_PATH}/csv_register/output_{today}.xlsx"
            print("\n")

            for idx, report in enumerate(Constants.CSV_REG_REPORT_TYPES):
                # if idx == 1:
                #     exit(0)

                print(
                    f"Processing HCM Excel column: {report['oracle_col']} && CSV Register Excel column: {report['csv_register_col']}")
                reportObj = GenerateReport(
                    oracle_col=report['oracle_col'],
                    adp_col=report['csv_register_col'],
                    output_sheet_name=report['output_sheet_name'],
                    output_filename=output_file_name,
                    adp_report=Constants.CSV_REG_REPORT

                )

                reportObj.generate_report()

                print(f"Sheet generated with sheet name: {report['output_sheet_name']} \n \n \n")

            print("CSV Register Recon Excel Successfully Generated.")

    if selection['Report'] == DAILY:

        # check if HCM Excel is present in the destination folder
        hcm_excel_is_present = GenerateReport.is_file_present(
            Constants.ORACLE_REPORT
        )

        # check if HCM Excel is present in the destination folder
        adp_excel_is_present = GenerateReport.is_file_present(
            Constants.ADP_REPORT
        )

        if hcm_excel_is_present and adp_excel_is_present:

            print(f"Next, your patient will be tested now. Please drink some water and relax \n")

            # create the output file with today's date'
            today = datetime.datetime.now().strftime('%m_%d_%Y')
            output_file_name = f"{Constants.OUTPUT_REPORT_PATH}/daily_report/output_{today}.xlsx"
            print("\n")

            for idx, report in enumerate(Constants.REPORT_TYPES):
                # if idx == 2:
                #     exit(0)

                print(f"Processing HCM Excel column: {report['oracle_col']} && ADP Excel column: {report['adp_col']}")

                reportObj = GenerateReport(
                    oracle_col=report['oracle_col'],
                    adp_col=report['adp_col'],
                    output_sheet_name=report['output_sheet_name'],
                    output_filename=output_file_name,
                    adp_report=Constants.ADP_REPORT
                )

                reportObj.generate_report()

                print(f"Sheet generated with sheet name: {report['output_sheet_name']} \n \n \n")

        print("Daily Recon Excel Successfully Generated.")
