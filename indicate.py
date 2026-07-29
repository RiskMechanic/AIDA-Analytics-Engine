import pandas as pd

df = pd.read_excel("output/result_raw.xlsx", sheet_name="ind", header=None)
header = df.iloc[0]
valid_cols = [col for col in header if isinstance(col, str) and col.startswith("31/12/")]
valid_cols = [header[0]] + valid_cols  # Include first column (indicator names)
df.columns = header
df = df[1:]  # Drop header row
df_cleaned = df[valid_cols]
df_cleaned.to_excel("output/indicators_cleaned.xlsx", index=False)
print("Cleaned indicator data saved.")

# first part of code cleans and saves the indicators sheet alone;
# below it renders graphs and insights based on cleaned data

