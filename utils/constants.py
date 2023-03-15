class Constants:
    ORACLE_REPORT = 'C:/Users/krana/Documents/FORTINET/Audit_Daily_RC/Output1.xlsx'
    ADP_REPORT = 'C:/Users/krana/Documents/FORTINET/Audit_Daily_RC/Deduction from Profile Updated.xlsx'
    AUDIT_REG_REPORT = 'C:/Users/krana/Documents/FORTINET/Audit_Daily_RC/Audit Register.xlsx'
    UNIQUE_COLUMN = "File_Number"
    V_LOOKUP_ACTION = "outer"
    OUTPUT_REPORT = './output.xlsx'
    OUTPUT_REPORT_PATH = 'C:/Users/krana/Documents/FORTINET/Audit_Daily_RC/reports'
    OUTPUT_AUDIT_REPORT_PATH = 'C:/Users/krana/Documents/FORTINET/Audit_Daily_RC/reports/audit_register'
    DELTA_COL = 'Delta'
    ADP_SHEET_NAME = '1'
    ORACLE_SHEET_NAME = 'Sheet1'
    EMPLOYEE_FIRST_NAME = "Employee_s_First_Name"
    EMPLOYEE_LAST_NAME = "Employee_Last_Name"
    PAYROLL_NAME = 'Payroll Name'
    REPORT_TYPES = [
        {
            "oracle_col": 'M_Pre_Tax_Medical_Deduction_Amount',
            "adp_col": 'M_Pre-Tax Medical_Deduction Amount',
            "output_sheet_name": 'M1',
        },
        {
            "oracle_col": 'V1_Pre_Tax_Vision_Deduction_Amount',
            "adp_col": 'V1_Pre-Tax Vision_Deduction Amount',
            "output_sheet_name": 'V1',
        },
        {
            "oracle_col": 'D1_Pre_Tax_Dental_Deduction_Amount',
            "adp_col": 'D1_Pre-Tax Dental_Deduction Amount',
            "output_sheet_name": 'D1',
        },
        {
            "oracle_col": 'M2_Post_Tax_Medica_Deduction_Amount',
            "adp_col": 'M2_Post Tax Medica_Deduction Amount',
            "output_sheet_name": 'M2',
        },
        {
            "oracle_col": 'V2_Post_Tax_Vision_Deduction_Amount',
            "adp_col": 'V2_Post Tax Vision_Deduction Amount',
            "output_sheet_name": 'V2',
        },
        {
            "oracle_col": 'D2_Post_Tax_Dental_Deduction_Amount',
            "adp_col": 'D2_Post Tax Dental_Deduction Amount',
            "output_sheet_name": 'D2',
        },
        {
            "oracle_col": '12_Supp_Life_EE_Deduction_Amount',
            "adp_col": '12_Supp Life - EE_Deduction Amount',
            "output_sheet_name": '12',
        },
        {
            "oracle_col": '13_Supp_Life_Sp_Deduction_Amount',
            "adp_col": '13_Supp Life -  Sp_Deduction Amount',
            "output_sheet_name": '13',
        },
        {
            "oracle_col": '_14_Supp_Life_Chi_Deduction_Amount',
            "adp_col": '14_Supp Life - Chi_Deduction Amount',
            "output_sheet_name": '14',
        },
        {
            "oracle_col": 'A5_PET_INSURANCE_Deduction_Amount',
            "adp_col": 'A5_PET INSURANCE_Deduction Amount',
            "output_sheet_name": 'A5',
        },
        {
            "oracle_col": 'A6_AUTO_HOME_RENT_Deduction_Amount',
            "adp_col": 'A6_AUTO/HOME/RENT_Deduction Amount',
            "output_sheet_name": 'A6',
        },
        {
            "oracle_col": 'H_Health_Fsa_Deduction_Amount',
            "adp_col": 'H_Health Fsa_Deduction Amount',
            "output_sheet_name": 'H',
        },
        {
            "oracle_col": 'C_Dep_Care_Fsa_Deduction_Amount',
            "adp_col": 'C_Dep Care Fsa_Deduction Amount',
            "output_sheet_name": 'C',
        },
        {
            "oracle_col": 'J_LPFSA_Deduction_Amount',
            "adp_col": 'J_LPFSA_Deduction Amount',
            "output_sheet_name": 'J',
        },
        {
            "oracle_col": 'I_Imputed_Medical_Deduction_Amount',
            "adp_col": 'I_Imputed Medical_Deduction Amount',
            "output_sheet_name": 'I',
        },
        {
            "oracle_col": 'D3_Imputed_inc_Dental_Deduction_Amount',
            "adp_col": 'D3_Imputed inc vis_Deduction Amount',
            "output_sheet_name": 'D3',
        },
        {
            "oracle_col": 'V3_Imputed_Vision_Deduction_Amount',
            "adp_col": 'V3_Imputed Vision_Deduction Amount',
            "output_sheet_name": 'V3',
        },
        {
            "oracle_col": 'LTD',
            "adp_col": 'LTD_LONG TERM DIS_Deduction Amount',
            "output_sheet_name": 'LTD',
        },
        {
            "oracle_col": 'ML_METLAW_Deduction_Amount',
            "adp_col": 'ML_METLAW_Deduction Amount',
            "output_sheet_name": 'ML',
        },
        {
            "oracle_col": 'HS_HSA_Deduction_Amount',
            "adp_col": 'HS_HSA SINGLE_Deductions',
            "output_sheet_name": 'HS',
        },
        {
            "oracle_col": 'HF_HSA_Deduction_Amount',
            "adp_col": 'HF_HSA FAMILY_Deductions',
            "output_sheet_name": 'HF',
        },
        {
            "oracle_col": 'HS_HSA_Deduction_Amount',
            "adp_col": 'HSS_HSA SINGLE_Deductions',
            "output_sheet_name": 'HSS',
        },
        {
            "oracle_col": 'HF_HSA_Deduction_Amount',
            "adp_col": 'HSF_HSA FAMILY_Deductions',
            "output_sheet_name": 'HSF',
        }

    ]
    AUDIT_REPORT_TYPES = [
        {
            "oracle_col": 'M_Pre_Tax_Medical_Deduction_Amount',
            "adp_register_col": 'DED CD M',
            "output_sheet_name": 'M1',
        },
        {
            "oracle_col": 'V1_Pre_Tax_Vision_Deduction_Amount',
            "adp_register_col": 'DED CD V1',
            "output_sheet_name": 'V1',
        },
        {
            "oracle_col": 'D1_Pre_Tax_Dental_Deduction_Amount',
            "adp_register_col": 'DED CD D1',
            "output_sheet_name": 'D1',
        },
        {
            "oracle_col": 'M2_Post_Tax_Medica_Deduction_Amount',
            "adp_register_col": 'DED CD M2',
            "output_sheet_name": 'M2',
        },
        {
            "oracle_col": 'V2_Post_Tax_Vision_Deduction_Amount',
            "adp_register_col": 'DED CD V2',
            "output_sheet_name": 'V2',
        },
        {
            "oracle_col": 'D2_Post_Tax_Dental_Deduction_Amount',
            "adp_register_col": 'DED CD D2',
            "output_sheet_name": 'D2',
        },
        {
            "oracle_col": '12_Supp_Life_EE_Deduction_Amount',
            "adp_register_col": 'DED CD 12',
            "output_sheet_name": '12',
        },
        {
            "oracle_col": '13_Supp_Life_Sp_Deduction_Amount',
            "adp_register_col": 'DED CD 13',
            "output_sheet_name": '13',
        },
        {
            "oracle_col": '_14_Supp_Life_Chi_Deduction_Amount',
            "adp_register_col": 'DED CD 14',
            "output_sheet_name": '14',
        },
        {
            "oracle_col": 'A5_PET_INSURANCE_Deduction_Amount',
            "adp_register_col": 'DED CD A5',
            "output_sheet_name": 'A5',
        },
        {
            "oracle_col": 'A6_AUTO_HOME_RENT_Deduction_Amount',
            "adp_register_col": 'DED CD A6',
            "output_sheet_name": 'A6',
        },
        {
            "oracle_col": 'H_Health_Fsa_Deduction_Amount',
            "adp_register_col": 'DED CD H',
            "output_sheet_name": 'H',
        },
        {
            "oracle_col": 'C_Dep_Care_Fsa_Deduction_Amount',
            "adp_register_col": 'DED CD C',
            "output_sheet_name": 'C',
        },
        {
            "oracle_col": 'J_LPFSA_Deduction_Amount',
            "adp_register_col": 'DED CD J',
            "output_sheet_name": 'J',
        },
        {
            "oracle_col": 'I_Imputed_Medical_Deduction_Amount',
            "adp_register_col": 'DED CD I',
            "output_sheet_name": 'I',
        },
        {
            "oracle_col": 'D3_Imputed_inc_Dental_Deduction_Amount',
            "adp_register_col": 'DED CD D3',
            "output_sheet_name": 'D3',
        },
        {
            "oracle_col": 'V3_Imputed_Vision_Deduction_Amount',
            "adp_register_col": 'DED CD V3',
            "output_sheet_name": 'V3',
        },
        {
            "oracle_col": 'LTD',
            "adp_register_col": 'DED CD LTD',
            "output_sheet_name": 'LTD',
        },
        {
            "oracle_col": 'ML_METLAW_Deduction_Amount',
            "adp_register_col": 'DED CD ML',
            "output_sheet_name": 'ML',
        },
        {
            "oracle_col": 'HS_HSA_Deduction_Amount',
            "adp_register_col": 'DED CD HS',
            "output_sheet_name": 'HS',
        },
        {
            "oracle_col": 'HF_HSA_Deduction_Amount',
            "adp_register_col": 'DED CD HF',
            "output_sheet_name": 'HF',
        },
        {
            "oracle_col": 'HS_HSA_Deduction_Amount',
            "adp_register_col": 'DED CD HSO',
            "output_sheet_name": 'HSS',
        },
        {
            "oracle_col": 'HF_HSA_Deduction_Amount',
            "adp_register_col": 'DED CD HFO',
            "output_sheet_name": 'HSF',
        }
    ]
