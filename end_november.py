import pandas as pd

# Reading the first Excel file
file_path_1 = 'data/12-инвест каталог sheet1.xlsx'
df1 = pd.read_excel(file_path_1)

# Reading the second Excel file
file_path_2 = 'data/table_data - 2025-11-28T171733.084.xlsx'
df2 = pd.read_excel(file_path_2)

# Displaying the column names of both files to ensure they've been read correctly
print("First Excel File (df1) Columns:")
print(df1.columns)

print("\nSecond Excel File (df2) Columns:")
print(df2.columns)
