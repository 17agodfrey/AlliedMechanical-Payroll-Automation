import tkinter as tk
from tkinter import filedialog, messagebox
from app import extract_job_costing_from_raw_excel

def select_payroll_file():
    filepath = filedialog.askopenfilename(
        title="Select Payroll Excel File",
        filetypes=[("Excel Files", "*.xls *.xlsx")]
    )
    if filepath:
        payroll_file_var.set(filepath)
        payroll_file_label.config(text=f"Payroll: {filepath.split('/')[-1]}")

def select_tax_file():
    filepath = filedialog.askopenfilename(
        title="Select Tax Excel File",
        filetypes=[("Excel Files", "*.xls *.xlsx")]
    )
    if filepath:
        tax_file_var.set(filepath)
        tax_file_label.config(text=f"Tax: {filepath.split('/')[-1]}")

def process_files():
    payroll_path = payroll_file_var.get()
    tax_path = tax_file_var.get()

    if not payroll_path or not tax_path:
        messagebox.showwarning("Missing Files", "Please select both payroll and tax files.")
        return

    output_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Save Output File As"
    )

    if not output_path:
        return

    try:
        extract_job_costing_from_raw_excel(payroll_path, tax_path, output_path)
        messagebox.showinfo("Success", f"File saved to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")

# GUI setup
root = tk.Tk()
root.title("Job Costing Formatter")
root.geometry("900x300")

payroll_file_var = tk.StringVar()
tax_file_var = tk.StringVar()

# Frame for file selectors
file_frame = tk.Frame(root)
file_frame.pack(pady=20)

font_large = ("Arial", 12)

# Payroll column
payroll_column = tk.Frame(file_frame)
payroll_column.grid(row=0, column=0, padx=(0, 100))  # Increased right padding for spacing

tk.Label(payroll_column, text="Payroll File", font=font_large).pack()
tk.Button(payroll_column, text="Browse Payroll", font=font_large, command=select_payroll_file).pack(pady=5)
payroll_file_label = tk.Label(payroll_column, text="No file selected", fg="gray", font=font_large)
payroll_file_label.pack()

# Tax column
tax_column = tk.Frame(file_frame)
tax_column.grid(row=0, column=1, padx=(100, 0))  # Increased left padding for spacing

tk.Label(tax_column, text="Tax File", font=font_large).pack()
tk.Button(tax_column, text="Browse Tax", font=font_large, command=select_tax_file).pack(pady=5)
tax_file_label = tk.Label(tax_column, text="No file selected", fg="gray", font=font_large)
tax_file_label.pack()

# Process button
tk.Button(root, text="Process and Save Output", command=process_files, bg="green", fg="white", font=("Arial", 14, "bold")).pack(pady=20)

root.mainloop()
