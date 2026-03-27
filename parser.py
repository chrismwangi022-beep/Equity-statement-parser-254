import pypdf
import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
from datetime import datetime
from dateutil.relativedelta import relativedelta

def get_analysis_window():
    """Calculates the 6-month window based on the 15th-day rule."""
    today = datetime.now()
    # If today is 15th or earlier, start from end of last month
    if today.day <= 15:
        end_date = today.replace(day=1) - relativedelta(days=1)
    else:
        end_date = today
    start_date = (end_date - relativedelta(months=5)).replace(day=1)
    return start_date, end_date

def parse_pdf(pdf_path, password=None):
    """Extracts text and uses Regex to find transactions."""
    transactions = []
    # Pattern specifically tuned for Equity Statement layout
    pattern = re.compile(r'(\d{2}-\d{2}-\d{4})\s+(\d{2}-\d{2}-\d{4})\s+(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})')

    try:
        reader = pypdf.PdfReader(pdf_path)
        if reader.is_encrypted:
            if not reader.decrypt(password):
                return None, "Invalid Password"

        for page in reader.pages:
            text = page.extract_text()
            lines = text.split('\n')
            for line in lines:
                match = pattern.search(line)
                if match:
                    t_date, v_date, desc, amt, bal = match.groups()
                    transactions.append({
                        'Date': datetime.strptime(t_date, '%d-%m-%Y'),
                        'Description': desc.strip(),
                        'Amount': float(amt.replace(',', '')),
                        'Balance': float(bal.replace(',', ''))
                    })
        
        if not transactions:
            return None, "No transactions found. Check if this is a standard Equity Statement."

        df = pd.DataFrame(transactions).sort_values('Date')
        
        # Determine Debit vs Credit based on Balance movement
        df['Prev_Balance'] = df['Balance'].shift(1)
        # If balance went up, it's a Credit. Otherwise, it's a Debit.
        df['Debit'] = df.apply(lambda x: x['Amount'] if x['Balance'] < x['Prev_Balance'] else 0, axis=1)
        df['Credit'] = df.apply(lambda x: x['Amount'] if x['Balance'] > x['Prev_Balance'] else 0, axis=1)
        
        return df, None

    except Exception as e:
        return None, str(e)

def run_app():
    root = tk.Tk()
    root.withdraw()

    # 1. Select File
    pdf_path = filedialog.askopenfilename(title="Select Equity Bank Statement", filetypes=[("PDF files", "*.pdf")])
    if not pdf_path: return

    # 2. Select Output Folder
    output_folder = filedialog.askdirectory(title="Select Where to Save Excel Report")
    if not output_folder: return

    # 3. Handle Password
    reader = pypdf.PdfReader(pdf_path)
    password = ""
    if reader.is_encrypted:
        password = simpledialog.askstring("Password", "Enter PDF Password:", show='*')
        if not password: return

    # 4. Process
    df, error = parse_pdf(pdf_path, password)
    
    if error:
        messagebox.showerror("Error", f"Failed to parse: {error}")
        return

    # 5. Apply 6-Month Analysis Logic
    start, end = get_analysis_window()
    mask = (df['Date'] >= start) & (df['Date'] <= end)
    df_6months = df.loc[mask].copy()
    
    # 6. Generate Summary
    df_6months['Month'] = df_6months['Date'].dt.strftime('%Y-%m')
    summary = df_6months.groupby('Month').agg({'Credit': 'sum', 'Debit': 'sum'})
    summary['Net Cashflow'] = summary['Credit'] - summary['Debit']

    # 7. Save to Excel
    try:
        output_path = os.path.join(output_folder, f"Equity_Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx")
        with pd.ExcelWriter(output_path) as writer:
            df.to_excel(writer, sheet_name='Full History', index=False)
            summary.to_excel(writer, sheet_name='6-Month Summary')
        
        messagebox.showinfo("Success", f"Analysis Complete!\nPeriod: {start.date()} to {end.date()}\nSaved to: {output_path}")
    except Exception as e:
        messagebox.showerror("Save Error", f"Could not save file: {e}")

if __name__ == "__main__":
    run_app()