import pandas as pd
import re
import os
from formatting import apply_formatting_to_excel

def extract_job_costing_from_raw_excel(filepath, output_file):
    df = pd.read_excel(filepath, sheet_name='Payroll Register', header=None)
    job_data = []
    current_employee = None
    
    job_lookup = {}
    total_bonus = 0.0  # NEW: to track total bonus
    un_coded_pay = 0.0
    gross_total = 0.0



    for i in range(len(df)):
        row = df.iloc[i]

        # Check for bonus in column 6 (7th column)
        if isinstance(row[6], str) and 'BN' in row[6]:
            match = re.search(r'BN\s*([\d,]+\.\d{2})', row[6])
            if match:
                bonus_value = float(match.group(1).replace(',', ''))
                total_bonus += bonus_value

        # Skip non-string rows in column 0
        if not isinstance(row[0], str):
            continue

        # Check if the row marks a new employee
        if "Associate ID" in row[0]:
            current_employee = row[0].split('\n')[0] if '\n' in row[0] else row[0]

            # If W-In Cost is in the same row, extract job number
            if "W-In Cost" in row[0]:
                match = re.search(r'W-In Cost:\s*(\d{4,})-(\d{4,})', row[0])
                job_number = match.group(2) if match else '0001'
            else:
                job_number = '0001'  # Default job number if missing

            # Try to extract pay even if W-In Cost is missing
            try:
                reg_pay = float(str(row[4]).replace(',', '')) if pd.notna(row[4]) else 0.0
                ot_pay = float(str(row[5]).replace(',', '')) if pd.notna(row[5]) else 0.0

                if reg_pay == 0.0 and ot_pay == 0.0:
                    ## skip this row
                    continue

                key = (current_employee, job_number)
                if key in job_lookup:
                    job_lookup[key]['Reg Pay'] += reg_pay
                    job_lookup[key]['OT Pay'] += ot_pay
                else:
                    job_lookup[key] = {
                        'Employee': current_employee,
                        'Job Number': job_number,
                        'Reg Pay': reg_pay,
                        'OT Pay': ot_pay
                    }
            except Exception as e:
                print(f"Error parsing row {i}: {e}")

        # Process follow-up rows for current employee if they contain W-In Cost
        elif "W-In Cost" in row[0] and current_employee:
            match = re.search(r'W-In Cost:\s*(\d{4,})-(\d{4,})', row[0])
            job_number = match.group(2) if match else '0001'

            try:
                reg_pay = float(str(row[4]).replace(',', '')) if pd.notna(row[4]) else 0.0
                ot_pay = float(str(row[5]).replace(',', '')) if pd.notna(row[5]) else 0.0

                key = (current_employee, job_number)
                if key in job_lookup:
                    job_lookup[key]['Reg Pay'] += reg_pay
                    job_lookup[key]['OT Pay'] += ot_pay
                else:
                    job_lookup[key] = {
                        'Employee': current_employee,
                        'Job Number': job_number,
                        'Reg Pay': reg_pay,
                        'OT Pay': ot_pay
                    }
            except Exception as e:
                print(f"Error parsing row {i}: {e}")
                
        elif "H Dept" in row[0] and current_employee: # Handle rows with no W-In Cost but containing Reg/OT pay
            try:
                reg_pay = float(str(row[4]).replace(',', '')) if pd.notna(row[4]) else 0.0
                ot_pay = float(str(row[5]).replace(',', '')) if pd.notna(row[5]) else 0.0
                if reg_pay > 0.0 or ot_pay > 0.0:
                    un_coded_pay += reg_pay + ot_pay
            except Exception as e:
                print(f"Error parsing uncoded pay row {i}: {e}")
                
        elif isinstance(row[7], str) and 'Gross' in row[7]:
            # Extract number after 'Gross'
            match = re.search(r'Gross\s*([\d,]+\.\d{2})', row[7])
            if match:
                gross_str = match.group(1).replace(',', '')
                try:
                    gross_val = float(gross_str)
                    gross_total += gross_val
                except Exception as e:
                    print(f"Error parsing Gross value on row {i}: {e}")        



    job_data = list(job_lookup.values())

    df_jobs = pd.DataFrame(job_data)
    job_numbers = sorted(df_jobs['Job Number'].unique())

    # Build column multi-index: (Job, Reg/O/T)
    columns = []
    for job in job_numbers:
        columns.append((job, 'Reg'))
        columns.append((job, 'O/T'))

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
            else:
                row.extend([None, None])
        employee_rows.append(row)
        employee_names.append(employee)

    output_df = pd.DataFrame(employee_rows, columns=multi_columns)
    output_df.insert(0, ('Job Number', ''), employee_names)

    # Build header rows
    file_title = os.path.splitext(os.path.basename(filepath))[0]
    header_row = [file_title]
    for job in job_numbers:
        header_row += ['Reg', 'O/T']

    top_row = ['Job Number']
    for job in job_numbers:
        top_row += [job, '']

    # Totals rows
    pay_total_row = ['pay total']
    ten_percent_row = ['.10']
    grand_total_row = ['TOTAL']

    total_pay = 0
    total_tax = 0

    for job in job_numbers:
        reg_col = pd.to_numeric(output_df[(job, 'Reg')], errors='coerce')
        ot_col = pd.to_numeric(output_df[(job, 'O/T')], errors='coerce')

        reg_total = reg_col.sum(skipna=True)
        ot_total = ot_col.sum(skipna=True)

        reg_10 = reg_total * 0.10
        ot_10 = ot_total * 0.10

        total_pay += reg_total + ot_total
        total_tax += reg_10 + ot_10

        pay_total_row.extend([round(reg_total, 2), round(ot_total, 2)])
        ten_percent_row.extend([round(reg_10, 2), round(ot_10, 2)])
        grand_total_row.extend([round(reg_total + reg_10, 2), round(ot_total + ot_10, 2)])

    # Append final values to each row
    pay_total_row.append(round(total_pay, 2))
    ten_percent_row.append(round(total_tax, 2))
    grand_total_row.append(round(total_pay + total_tax, 2))


    # Combine everything
    final_data = [top_row, header_row] + output_df.values.tolist() + [
        pay_total_row,
        ten_percent_row,
        grand_total_row
    ]

    final_df = pd.DataFrame(final_data)
    
    # === Add 3 empty rows and the summary block ===
    final_df = pd.concat([
        final_df,
        pd.DataFrame([[]]*3),  # 3 empty rows
        pd.DataFrame([
            [round(total_pay, 2)],
            [round(total_bonus, 2)],
            [round(un_coded_pay, 2)],
            [round(total_pay + total_bonus + un_coded_pay, 2)],
            [round(gross_total, 2)]  # New 5th summary row for gross pay
        ])
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
# extract_job_costing_from_raw_excel("payroll_test.xls", "formatted_output.xlsx")
