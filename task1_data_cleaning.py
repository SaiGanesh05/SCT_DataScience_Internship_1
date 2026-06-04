import pandas as pd
print("Script Started")
# Load Excel file
df = pd.read_excel(
    "Superstore_Sales_Dashboard.xlsx",
    sheet_name="Raw_Data"
)

print("Original Shape:", df.shape)

# Missing values before cleaning
print("\nMissing Values:")
print(df.isnull().sum())

# Handle missing values safely
for col in df.columns:

    # Numeric columns
    if pd.api.types.is_numeric_dtype(df[col]):
        df[col] = df[col].fillna(df[col].median())

    # Date columns
    elif "date" in col.lower():
        df[col] = pd.to_datetime(df[col], errors="coerce")

    # Text columns
    else:
        df[col] = df[col].fillna("Unknown")

# Remove duplicate rows
duplicates = df.duplicated().sum()
print("\nDuplicates Found:", duplicates)

df = df.drop_duplicates()

# Convert date columns
for col in df.columns:
    if "date" in col.lower():
        df[col] = pd.to_datetime(df[col], errors="coerce")

print("\nData Types:")
print(df.dtypes)

# Save cleaned data
df.to_csv("Superstore_Cleaned.csv", index=False)

print("\n✅ Data Cleaning Completed Successfully")
print("✅ Output File: Superstore_Cleaned.csv")
print("Script Finished")