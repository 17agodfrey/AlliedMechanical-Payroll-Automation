import re
import xlwings as xw

def calculate_tax(filepath):
    errors = []
    tax_data = {}

    # Open workbook using xlwings to get evaluated values
    app = xw.App(visible=False)
    try:
        wb = app.books.open(filepath)
        sheet = wb.sheets["Payroll History"]
        
        # Read all values into a 2D list
        data = sheet.used_range.value

        # Locate the "TOTAL EMPLOYER TAX" column
        tax_col_idx = None
        memo_col_idx = None
        header_row = data[0]
        for col_idx, value in enumerate(header_row):
            if isinstance(value, str):
                val_upper = value.upper()
                if "TOTAL EMPLOYER TAX" in val_upper:
                    tax_col_idx = col_idx
                elif "MEMO : K-401K MATCH" in val_upper:
                    memo_col_idx = col_idx

        if tax_col_idx is None:
            errors.append("Could not find 'TOTAL EMPLOYER TAX' column in the first row.")
            raise ValueError(errors[-1])

        if memo_col_idx is None:
            errors.append("Could not find 'MEMO : K-401K MATCH' column in the first row.")
            raise ValueError(errors[-1])

        for row in data[1:]:  # Skip header row
            if len(row) < max(6, tax_col_idx + 1, memo_col_idx + 1):
                continue

            raw_name = row[1]  # Column B
            try:
                dist = int(row[5])  # Column F
            except (ValueError, TypeError):
                dist = 0

            tax = row[tax_col_idx]
            memo_401k = row[memo_col_idx]

            if raw_name:
                name = re.sub(r'\s*,\s*', ', ', str(raw_name).strip())
                if name not in tax_data:
                    tax_data[name] = []
                tax_data[name].append((dist, tax, memo_401k))
    finally:
        wb.close()
        app.quit()

    return tax_data, errors


def get_tax(tax_data, employee_name, index):
    errors = []
    if employee_name in tax_data and index <= len(tax_data[employee_name]):
        for dist, emp_tax, memo_401k in tax_data[employee_name]:
            if dist == index:
                return emp_tax, memo_401k, errors 
        return 0.0, None, errors
    else: 
        errors.append(f"Error getting tax for Employee {employee_name}. Not found or index out of range.")
        return 0.0, None, errors
