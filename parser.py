import pypdf
import re
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

def parse_equity_to_cashflow(pdf_path, password=None):
    reader = pypdf.PdfReader(pdf_path)
    if reader.is_encrypted:
        reader.decrypt(password)
        
    transactions = []
    # Pattern: Date1 Date2 Description... Amount Balance
    pattern = re.compile(r'(\d{2}-\d{2}-\d{4})\s+(\d{2}-\d{2}-\d{4})\s+(.*?)\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})')

    for page in reader.pages:
        text = page.extract_text()
        lines = text.split('\n')
        for line in lines:
            match = pattern.search(line)
            if match:
                t_date, _, desc, amt, bal = match.groups()
                transactions.append({
                    'Date': datetime.strptime(t_date, '%d-%m-%Y'),
                    'Description': desc.strip(),
                    'Amount': float(amt.replace(',', '')),
                    'Balance': float(bal.replace(',', ''))
                })

    df = pd.DataFrame(transactions)
    
    # Apply Balance Logic to separate Debit/Credit
    df['Prev_Balance'] = df['Balance'].shift(1)
    df['Type'] = df.apply(lambda x: 'Credit' if x['Balance'] > x['Prev_Balance'] else 'Debit', axis=1)
    # Fix the first row manually or based on opening balance
    
    return df

def get_6_month_summary(df):
    today = datetime.now()
    # 15th Day Rule
    end_date = today.replace(day=1) - relativedelta(days=1) if today.day <= 15 else today
    start_date = (end_date - relativedelta(months=5)).replace(day=1)
    
    # Filter
    mask = (df['Date'] >= start_date) & (df['Date'] <= end_date)
    filtered_df = df.loc[mask].copy()
    filtered_df['Month'] = filtered_df['Date'].dt.strftime('%Y-%m')
    
    summary = filtered_df.groupby(['Month', 'Type'])['Amount'].sum().unstack(fill_value=0)
    summary['Net Surplus'] = summary.get('Credit', 0) - summary.get('Debit', 0)
    return summary

# Usage:
# df = parse_equity_to_cashflow('statement.pdf', 'password123')
# report = get_6_month_summary(df)
# report.to_excel('Cashflow_Summary.xlsx')