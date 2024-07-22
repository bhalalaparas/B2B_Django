import pandas as pd

file_path = '46e4be34-7d66-4381-8904-6fa7fba06912_1712915192000.xlsx'

df = pd.read_excel(file_path, sheet_name='Sales Report')

df_f = pd.DataFrame(df)
# print(df)
results = pd.DataFrame()
            
states = df_f["Customer's Billing State"].unique()
# print(states)
for state in states:
    df_a = df[df_f["Customer's Billing State"] == state]

    total_ap = df_a.groupby("Event Type")[[
        "Final Invoice Amount (Price after discount+Shipping Charges)", 
        "Taxable Value (Final Invoice Amount -Taxes)", 
        "IGST Amount", 
        "CGST Amount", 
        "SGST Amount (Or UTGST as applicable)"
    ]].sum().reset_index()
    total_ap["Customer's Billing State"] = state
    results = pd.concat([results, total_ap], ignore_index=True)

print(results)