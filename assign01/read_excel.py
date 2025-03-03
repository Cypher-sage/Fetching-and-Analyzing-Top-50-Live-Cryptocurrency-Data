import pandas as pd

excel_file = 'Top 50 Live Cryptocurrency Data.xlsx'
df = pd.read_excel(excel_file)

print("\nCryptocurrency Data from Excel:")
print("=" * 50)
print(df)

print("\nFirst 5 cryptocurrencies:")
print("=" * 50)
print(df.head())

print("\nSummary Statistics:")
print("=" * 50)
print(df.describe()) 