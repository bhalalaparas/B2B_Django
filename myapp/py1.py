import pandas as pd
import os
# from django.core.management.base import BaseCommand # type: ignore
# import getpass

# class Command(BaseCommand):
#     help = 'Run the server with a password prompt'

#     def handle(self, *args, **kwargs):
#         password = getpass.getpass(prompt='Enter Password: ')
#         if password == 'your_password_here':
#             self.stdout.write("Access Granted")
#             # Place the command to run server here
#             os.system('python manage.py runserver')
#         else:
#             self.stdout.write("Access Denied")


file1 = '042024_24BKLPV1834R1ZC_GSTR2B_20052024.xlsx'
file2 = 'DayBook.xlsx'

df2 = pd.read_excel(file1, sheet_name='B2B', header=[4, 5])
# print(df2)
df1 = pd.read_excel(file2, sheet_name='Purchase Register', header=6)
# print(df1)


# Flatten the multi-level columns in df2
df2.columns = [' '.join(col).strip() for col in df2.columns.values]

# print("Columns in df2 after flattening:")
# print(df2.columns)
# print("\nFirst few rows of df2:")
# print(df2.head())

# Select relevant columns
b_df = pd.DataFrame(df2, columns=[
    'GSTIN of supplier', 'Invoice number', 'Invoice Date', 'Invoice Value(₹)',
    'Central Tax(₹)', 'State/UT Tax(₹)'
])
# print(b_df)

df1.columns = df1.columns.str.strip()
c_df = pd.DataFrame(df1, columns=[
    'Supplier Invoice No.', 'GSTIN/UIN', 'Value', 'Gross Total', 'CGST', 'SGST', 'Supplier Invoice Date'
])
# print(c_df)

c_df['Supplier Invoice Date'] = pd.to_datetime(c_df['Supplier Invoice Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%Y-%m-%d')
b_df['Invoice Date'] = pd.to_datetime(b_df['Invoice Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%Y-%m-%d')

missing_in_daybook = []
missing_in_B2B = []
difference = []

c_df['Key'] = c_df['Supplier Invoice No.'] + '_' + c_df['GSTIN/UIN']
# print(c_df)
b_df['Invoice number'] = b_df['Invoice number'].astype(str)
b_df['Key'] = b_df['Invoice number'] + '_' + b_df['GSTIN of supplier']
# print(b_df)

# Debugging print statements
# print("c_df head:")
# print(c_df.head())
# print("b_df head:")
# print(b_df.head())

for index1, row1 in c_df.iterrows():
    match_found = False
    for index2, row2 in b_df.iterrows():
        if row1['Key'] == row2['Key']:
            match_found = True
            if row1['Gross Total'] != row2['Invoice Value']:
                difference.append(f"Different value between Invoice Value and Gross Total for {row1['Invoice Number']} (GSTIN: {row1['GSTIN']}): {row1['Gross Total']} != {row2['Invoice Value']}")
            if row1['CGST'] != row2['Central Tax']:
                difference.append(f"Different value between CGST and GST file's CGST for {row1['Invoice Number']} (GSTIN: {row1['GSTIN']}): {row1['CGST']} != {row2['Central Tax']}")
            if row1['SGST'] != row2['State Tax']:
                difference.append(f"Different value between SGST and GST file's SGST for {row1['Invoice Number']} (GSTIN: {row1['GSTIN']}): {row1['SGST']} != {row2['State Tax']}")
            if row1['Invoice Date'] != row2['Invoice Date']:
                difference.append(f"Different Invoice Dates for {row1['Invoice Number']} (GSTIN: {row1['GSTIN']}): {row1['Invoice Date']} != {row2['Invoice Date']}")
            break
    if not match_found:
        missing_in_B2B.append(f"{row1['Supplier Invoice No.']} (GSTIN/UIN: {row1['GSTIN/UIN']})")

for index2, row2 in b_df.iterrows():
    if row2['Key'] not in c_df['Key'].values:
        missing_in_daybook.append(f"{row2['Invoice number']} (GSTIN: {row2['GSTIN of supplier']})")

# Debugging print statements
# print("Discrepancies found:")
# print(difference)
# print("Missing in B2B:")
# print(missing_in_B2B)
# print("Missing in DayBook:")
# print(missing_in_daybook)

desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
output_file_path = os.path.join(desktop_path, 'output.txt')

with open(output_file_path, 'w', encoding='utf-8') as f:
    if difference:
        f.write("difference:\n")
        for line in difference:
            f.write(line + '\n')
        f.write('\n')
        
    if missing_in_daybook:
        f.write("Invoices present in GSTR2B but missing in DayBook:\n")
        for invoice in missing_in_daybook:
            f.write(str(invoice) + '\n') 
        f.write('\n')

    if missing_in_B2B:
        f.write("Invoices present in DayBook but missing in GSTR2B:\n")
        for invoice in missing_in_B2B:
            f.write(str(invoice) + '\n') 
        f.write('\n')

print("Output saved to desktop as output.txt")
