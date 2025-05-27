import re
import pandas as pd

def calculate_tax(filepath):
    errors = []
    
    # Read the whole sheet without assuming headers
    df = pd.read_excel(filepath, sheet_name="Payroll History", header=None)

    # Locate the "TOTAL EMPLOYER TAX" column in the first row
    tax_col_idx = None
    for col_idx, value in enumerate(df.iloc[0]):
        if isinstance(value, str) and "TOTAL EMPLOYER TAX" in value.upper():
            tax_col_idx = col_idx
            break
        
    if tax_col_idx is None:
        errors.append("Could not find 'TOTAL EMPLOYER TAX' column in the first row.")
        raise ValueError("Could not find 'TOTAL EMPLOYER TAX' column in the first row.")
        
    # Locate the "MEMO : K-401K MATCH" column in the first row
    memo_col_idx = None
    for col_idx, value in enumerate(df.iloc[0]):
        if isinstance(value, str) and "MEMO : K-401K MATCH" in value.upper():
            memo_col_idx = col_idx
            break
        
    if memo_col_idx is None:
        errors.append("Could not find 'MEMO : K-401K MATCH' column in the first row.")
        raise ValueError("Could not find 'MEMO : K-401K MATCH' column in the first row.")

    tax_data = {}

    for i, row in df.iterrows():
        raw_name = row[1]  # Column B
        
        try:
            dist = int(row[5])  # Column F
        except (ValueError, TypeError):
            dist = 0
        
        tax = row[tax_col_idx]  # Column with "TOTAL EMPLOYER TAX"
        memo_401k = row[memo_col_idx] # Column with "MEMO : K-401K MATCH"

        if pd.notna(raw_name):
            # Normalize name: remove extra spaces and enforce "<Last>, <First>" format with single space
            name = re.sub(r'\s*,\s*', ', ', str(raw_name).strip())
            
            if name not in tax_data:
                tax_data[name] = []
            tax_data[name].append((dist, tax, memo_401k))
    return tax_data, errors
            
def get_tax(tax_data, employee_name, index):
    errors = []
    if employee_name in tax_data and index <= len(tax_data[employee_name]):
        # go through the array of tuples and find the one with the same dist
        for dist, emp_tax, memo_401k in tax_data[employee_name]:
            if dist == index:
                return emp_tax, memo_401k, errors 
        return 0.0, errors
    else: 
        errors.append(f"Error getting tax for Employee {employee_name}. not found or index out of range.")
        # print(f"Employee {employee_name} not found or index out of range.")
        return 0.0, errors