import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import os

def run_converter():
    # 1. Initialize Tkinter and hide the main window
    root = tk.Tk()
    root.withdraw()

    # 2. Ask user to select the PDF file
    pdf_path = filedialog.askopenfilename(
        title="Select Equity Bank PDF Statement",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not pdf_path: return

    # 3. Ask for the Output Folder
    output_folder = filedialog.askdirectory(title="Select Folder to Save Excel File")
    if not output_folder: return

    # 4. Handle Password Protection
    password = None
    try:
        # Try opening without password first
        with pdfplumber.open(pdf_path) as pdf:
            pass
    except:
        # If it fails, ask the user for the password
        password = simpledialog.askstring("Password Required", "Enter PDF Password:", show='*')
        if not password:
            messagebox.showerror("Error", "Password is required to open this file.")
            return

    # 5. Extraction Logic
    try:
        all_data = []
        with pdfplumber.open(pdf_path, password=password) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    df_page = pd.DataFrame(table[1:], columns=table[0])
                    all_data.append(df_page)
        
        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            
            # Save to Excel
            output_name = os.path.join(output_folder, "Equity_Analysis_Output.xlsx")
            final_df.to_excel(output_name, index=False)
            messagebox.showinfo("Success", f"File saved successfully to:\n{output_name}")
        else:
            messagebox.showwarning("No Data", "Could not find any tables in the PDF.")
            
    except Exception as e:
        messagebox.showerror("Processing Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    run_converter()