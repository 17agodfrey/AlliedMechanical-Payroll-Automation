import pandas as pd
import re
import os
from formatting import apply_formatting_to_excel
from employer_tax import calculate_employer_tax, get_emp_tax

def update_bonus_vacation(cell_value, total_bonus, total_vacation):
    if isinstance(cell_value, str):
        if 'BN' in cell_value:
            match = re.search(r'BN\s*([\d,]+\.\d{2})', cell_value)
            if match:
                total_bonus += float(match.group(1).replace(',', ''))
        elif 'VAC' in cell_value:
            match = re.search(r'VAC\s*([\d,]+\.\d{2})', cell_value)
            if match:
                total_vacation += float(match.group(1).replace(',', ''))
    return total_bonus, total_vacation

def process_payroll_file(filepath, tax_data):
    df = pd.read_excel(filepath, sheet_name='Payroll Register', header=None)
    job_lookup = {}
    current_employee = None
    c_employee_row_i = 0;

    total_bonus = 0.0
    total_vacation = 0.0
    un_coded_pay = 0.0
    gross_total = 0.0

    for i in range(len(df)):
        row = df.iloc[i]                

        if not isinstance(row[0], str):
            continue

        if "Associate ID" in row[0]:
            current_employee = row[0].split('\n')[0] if '\n' in row[0] else row[0]
            current_employee = re.sub(r'\s*,\s*', ', ', str(current_employee).strip()) ## Normalize name, just in case 
            c_employee_row_i = 1
            
            ##handle bonus and vacation
            total_bonus, total_vacation = update_bonus_vacation(row[6], total_bonus, total_vacation)
            
            job_number = '0001'
            if "W-In Cost" in row[0]:
                match = re.search(r'W-In Cost:\s*(\d{4,})-(\d{4,})', row[0])
                job_number = match.group(2) if match else '0001'

            try:
                reg_pay = float(str(row[4]).replace(',', '')) if pd.notna(row[4]) else 0.0
                ot_pay = float(str(row[5]).replace(',', '')) if pd.notna(row[5]) else 0.0
                if reg_pay == 0.0 and ot_pay == 0.0:
                    if isinstance(row[3], str) and 'UTO' in row[3]:
                        c_employee_row_i = c_employee_row_i - 1
                    continue

                emp_tax = get_emp_tax(tax_data, current_employee, c_employee_row_i)

                key = (current_employee, job_number)
                if key in job_lookup:
                    job_lookup[key]['Reg Pay'] += reg_pay
                    job_lookup[key]['OT Pay'] += ot_pay
                    job_lookup[key]['Emp Tax'] += emp_tax
                else:
                    job_lookup[key] = {
                        'Employee': current_employee,
                        'Job Number': job_number,
                        'Reg Pay': reg_pay,
                        'OT Pay': ot_pay,
                        'Emp Tax': emp_tax
                    }
            except Exception as e:
                print(f"Error parsing row {i}: {e}")

        elif "W-In Cost" in row[0] and current_employee:
            match = re.search(r'W-In Cost:\s*(\d{4,})-(\d{4,})', row[0])
            job_number = match.group(2) if match else '0001'
            
            c_employee_row_i = c_employee_row_i + 1

            total_bonus, total_vacation = update_bonus_vacation(row[6], total_bonus, total_vacation)
            
            try:
                reg_pay = float(str(row[4]).replace(',', '')) if pd.notna(row[4]) else 0.0
                ot_pay = float(str(row[5]).replace(',', '')) if pd.notna(row[5]) else 0.0
                if reg_pay == 0.0 and ot_pay == 0.0:
                    continue
                
                emp_tax = get_emp_tax(tax_data, current_employee, c_employee_row_i)

                key = (current_employee, job_number)
                if key in job_lookup:
                    job_lookup[key]['Reg Pay'] += reg_pay
                    job_lookup[key]['OT Pay'] += ot_pay
                    job_lookup[key]['Emp Tax'] += emp_tax
                else:
                    job_lookup[key] = {
                        'Employee': current_employee,
                        'Job Number': job_number,
                        'Reg Pay': reg_pay,
                        'OT Pay': ot_pay,
                        'Emp Tax': emp_tax
                    }
            except Exception as e:
                print(f"Error parsing row {i}: {e}")
                            #handle UTO
        elif isinstance(row[3], str) and 'UTO' in row[3]:
                continue
        # Handle rows with no W-In Cost but containing Reg/OT pay (Uncoded pay, but not bonues and etc) 
        elif "H Dept" in row[0] and current_employee:              
            c_employee_row_i = c_employee_row_i + 1
            
            total_bonus, total_vacation = update_bonus_vacation(row[6], total_bonus, total_vacation)
            
            try:
                reg_pay = float(str(row[4]).replace(',', '')) if pd.notna(row[4]) else 0.0
                ot_pay = float(str(row[5]).replace(',', '')) if pd.notna(row[5]) else 0.0
                if reg_pay > 0.0 or ot_pay > 0.0:
                    un_coded_pay += reg_pay + ot_pay
            except Exception as e:
                print(f"Error parsing uncoded pay row {i}: {e}")

        elif isinstance(row[7], str) and 'Gross' in row[7]:
            match = re.search(r'Gross\s*([\d,]+\.\d{2})', row[7])
            if match:
                try:
                    gross_total += float(match.group(1).replace(',', ''))
                except Exception as e:
                    print(f"Error parsing Gross value on row {i}: {e}")
    

    ######################################################################## 

    job_data = list(job_lookup.values())

    df_jobs = pd.DataFrame(job_data)
    job_numbers = sorted(df_jobs['Job Number'].unique())

    # Build column multi-index: (Job, Reg/O/T)
    columns = []
    for job in job_numbers:
        columns.append((job, 'Reg'))
        columns.append((job, 'O/T'))
        columns.append((job, 'Emp Tax'))

    multi_columns = pd.MultiIndex.from_tuples(columns)

    # Build data rows for each employee
    employee_rows = []
    employee_names = []
    for employee in df_jobs['Employee'].unique():
        emp_jobs = df_jobs[df_jobs['Employee'] == employee]
        row = []
        for job in job_numbers:
            job_entry = emp_jobs[emp_jobs['Job Number'] == job]
            if not job_entry.empty:
                row.append(job_entry['Reg Pay'].values[0])
                row.append(job_entry['OT Pay'].values[0])
                row.append(job_entry['Emp Tax'].values[0])
            else:
                row.extend([None, None, None])
        employee_rows.append(row)
        employee_names.append(employee)

    output_df = pd.DataFrame(employee_rows, columns=multi_columns)
    output_df.insert(0, ('Job Number', ''), employee_names)

    # Build header rows
    file_title = os.path.splitext(os.path.basename(filepath))[0]
    header_row = [file_title]
    for job in job_numbers:
        header_row += ['Reg', 'O/T', 'Emp Tax']

    top_row = ['Job Number']
    for job in job_numbers:
        top_row += [job, '', '']

    # Totals rows
    pay_total_row = ['pay total']
    grand_total_row = ['TOTAL']

    grand_sum_total_pay = 0
    grand_sum_total_tax = 0
    
    total_pay_per_job = {}
    total_emp_tax_per_job = {}

    for job in job_numbers:
        reg_col = pd.to_numeric(output_df[(job, 'Reg')], errors='coerce')
        ot_col = pd.to_numeric(output_df[(job, 'O/T')], errors='coerce')
        emp_tax_col = pd.to_numeric(output_df[(job, 'Emp Tax')], errors='coerce')

        reg_total = reg_col.sum(skipna=True)
        ot_total = ot_col.sum(skipna=True)
        emp_tax_total = emp_tax_col.sum(skipna=True)

        grand_sum_total_pay += reg_total + ot_total
        grand_sum_total_tax += emp_tax_total
        
        total_pay_per_job[job] = reg_total + ot_total
        total_emp_tax_per_job[job] = emp_tax_total

        pay_total_row.extend([round(reg_total, 2), round(ot_total, 2), round(emp_tax_total, 2)])
        grand_total_row.extend([round(reg_total + ot_total + emp_tax_total, 2), '', ''])

    # Append final values to each row ------ idk what this was doing tbh
    pay_total_row.append(round(grand_sum_total_pay, 2))
    grand_total_row.append(round(grand_sum_total_pay + grand_sum_total_tax, 2))


    # Combine everything
    final_data = [top_row, header_row] + output_df.values.tolist() + [
        pay_total_row,
        grand_total_row
    ]

    final_df = pd.DataFrame(final_data)
    return final_df, grand_sum_total_pay, total_bonus, total_vacation, un_coded_pay, gross_total, total_pay_per_job, total_emp_tax_per_job


def extract_job_costing_from_raw_excel(payroll_filepath, tax_filepath, output_file):
    
    # Process the employer tax file
    tax_data = calculate_employer_tax(tax_filepath)
    
    
    # Process the payroll file
    payroll_df, total_pay, total_bonus, total_vacation, un_coded_pay, gross_total, total_pay_per_job, total_emp_tax_per_job = process_payroll_file(payroll_filepath, tax_data)
    
    # Get all unique job numbers from both dictionaries
    all_jobs = sorted(set(total_pay_per_job.keys()).union(total_emp_tax_per_job.keys()))

    # Create summary rows: Job Number, Pay, Emp Tax, Total Cost
    summary_rows = []
    for job in all_jobs:
        pay = round(total_pay_per_job.get(job, 0), 2)
        emp_tax = round(total_emp_tax_per_job.get(job, 0), 2)
        total_cost = round(pay + emp_tax, 2)
        summary_rows.append([job, pay, emp_tax, total_cost])

    # Convert to DataFrame
    summary_df = pd.DataFrame(summary_rows)
    
    # === Add 3 empty rows and the summary block ===
    final_df = pd.concat([
        payroll_df,
        pd.DataFrame([[]]*3),  # 3 empty rows
        pd.DataFrame([
            [round(total_pay, 2), 'Total Pay (Reg + O/T)'],
            [round(total_bonus, 2), 'Total Bonus'],
            [round(total_vacation, 2), 'Total Vacation'],
            [round(un_coded_pay, 2), 'Total Uncoded Pay'],
            [round(total_pay + total_bonus + total_vacation + un_coded_pay, 2), 'Sheet Total Pay + Bonus + Vacation + Uncoded'],
            [round(gross_total, 2), 'Report Total (Gross)'],  # New 5th summary row for gross pay
            [abs(round(gross_total - (total_pay + total_bonus + total_vacation + un_coded_pay), 2)), 'diff'],  # New 6th summary row for gross pay minus bonus
            
        ]),
        pd.DataFrame([[]]*3),  # 3 empty rows
        pd.DataFrame([['Job Number', 'Pay (Reg + O/T)', 'Emp Tax', 'Total Job Cost']]),  # Header for the summary block        
        summary_df
        # # go through total pay per job, and add them to the end of the dataframe
        # pd.DataFrame([[job, round(total_pay_per_job[job], 2)] for job in total_pay_per_job.keys(), ]),
    ], ignore_index=True)

    # Write to Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, header=False)

    # Set first column width to 165 pixels (~22.5 Excel width units)
    from openpyxl import load_workbook
    wb = load_workbook(output_file)
    ws = wb.active
    ws.column_dimensions['A'].width = 25  # 1 Excel width unit ≈ 7 pixels
    wb.save(output_file)

    apply_formatting_to_excel(output_file)

    return output_file

# Example:
# extract_job_costing_from_raw_excel("payroll/payroll_test_12.xls", "tax/tax_test_12.xlsx", "formatted_output_12.xlsx")
# extract_job_costing_from_raw_excel("payroll/payroll_test_16.xls", "tax/tax_test_16.xlsx", "formatted_output_16.xlsx")
# extract_job_costing_from_raw_excel("payroll/payroll_test_18.xls", "tax/tax_test_18.xlsx", "formatted_output_18.xlsx")
# extract_job_costing_from_raw_excel("payroll/payroll_test_20.xls", "tax/tax_test_20.xlsx", "formatted_output_20.xlsx")
