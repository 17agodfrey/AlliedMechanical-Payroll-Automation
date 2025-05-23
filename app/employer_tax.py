import re
import pandas as pd

def calculate_employer_tax(filepath):
    # Read the whole sheet without assuming headers
    df = pd.read_excel(filepath, sheet_name="Payroll History", header=None)

    # Locate the "TOTAL EMPLOYER TAX" column in the first row
    tax_col_idx = None
    for col_idx, value in enumerate(df.iloc[0]):
        if isinstance(value, str) and "TOTAL EMPLOYER TAX" in value.upper():
            tax_col_idx = col_idx
            break

    if tax_col_idx is None:
        raise ValueError("Could not find 'TOTAL EMPLOYER TAX' column in the first row.")

    tax_data = {}

    for i, row in df.iterrows():
        raw_name = row[1]  # Column B
        
        try:
            dist = int(row[5])  # Column F
        except (ValueError, TypeError):
            dist = 0
        
        tax = row[tax_col_idx]  # Column AG

        if pd.notna(raw_name):
            # Normalize name: remove extra spaces and enforce "<Last>, <First>" format with single space
            name = re.sub(r'\s*,\s*', ', ', str(raw_name).strip())
            
            if name not in tax_data:
                tax_data[name] = []
            tax_data[name].append((dist, tax))
    return tax_data
            
def get_emp_tax(tax_data, employee_name, index):
    if employee_name in tax_data and index <= len(tax_data[employee_name]):
        # go through the array of tuples and find the one with the same dist
        for dist, tax in tax_data[employee_name]:
            if dist == index:
                return tax
        return 0.0
    else: 
        print(f"Employee {employee_name} not found or index out of range.")
        return 0.0