import pandas as pd
import openpyxl

df = pd.read_excel('test_invoice.xlsx')
df = df.iloc[0:7, 0:6]
print(df.head(10))