# Freezing Dates and Batch Summary from Excel

This Google Colab notebook processes an Excel file containing product stock data to extract unique batch numbers, slaughter dates, production/debond dates, and expiry dates. It formats these dates and values to comply with Spanish health authority requirements, ensuring that each unique value is followed by a comma for clear separation in certification reports.

## Purpose

The notebook is designed to be used directly in Google Colab so all team members can run it without installing any software. Users only need to upload the Excel file provided by the company's stock management program.

## Features

- Uploads an Excel file (.xlsx) directly in Colab.
- Reads the file skipping initial rows to reach the relevant data.
- Cleans and standardizes column names.
- Extracts and groups data by product name, ensuring unique values.
- Formats dates in DD/MM/YYYY format.
- Adds a comma after each unique value to facilitate distinction in official certificates.
- Displays a summary in the console.
- Exports the summarized data as an Excel file for download.

## How to Use

1. Open the notebook in Google Colab.
2. Run the cells sequentially.
3. Upload the Excel file when prompted.
4. Review the printed summary in the console.
5. Download the generated `resumen_por_nombre.xlsx` file.

## Code Overview

The core logic of the notebook includes:

- Reading and cleaning the Excel data.
- Renaming columns for easier reference.
- Filtering invalid or empty product names.
- Formatting batch numbers and dates.
- Grouping data by product and extracting unique values with commas.
- Exporting the final summary to an Excel file.

## Requirements

- Run in Google Colab environment.
- Requires `pandas` library (pre-installed in Colab).
- Input file should be an Excel `.xlsx` matching the stock program export.

## Author

dominberbel98

---

## Code

```python
import pandas as pd
from google.colab import files

# Step 1: Upload the Excel file
print("Please upload the Excel file:")
uploaded = files.upload()
file = list(uploaded.keys())[0]

# Step 2: Read Excel skipping first 8 rows, header is on row 10 (index 9)
df = pd.read_excel(file, skiprows=8, header=1)

# Step 3: Clean column names
df.columns = df.columns.str.replace('\n', ' ').str.strip()

# Step 4: Rename key columns for easier access
df = df.rename(columns={
    'Nombre Name': 'Nombre',
    'Lote Batch': 'Lote',
    'F.Sacrificio Slaughter D.': 'F_Sacrificio',
    'D.prod/desp Prod/Debond': 'D_Prod_Desp',
    'F.Caducidad Expiry Date': 'F_Caducidad'
})

# Step 5: Clean and standardize "Nombre"
df['Nombre'] = df['Nombre'].astype(str).str.strip().str.upper()

# Remove rows with invalid or empty names
df = df[~df['Nombre'].isin(['', 'NAN', 'nan', 'NaN']) & df['Nombre'].notna()]

# Ensure batch numbers are strings and remove trailing ".0" if any
df['Lote'] = df['Lote'].astype(str).str.replace('.0$', '', regex=True)

# Step 6: Convert date columns to DD/MM/YYYY format
for col in ['F_Sacrificio', 'D_Prod_Desp', 'F_Caducidad']:
    df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%d/%m/%Y')

# Step 7: Group by "Nombre" and extract unique values with trailing commas
result = {}

for nombre, group in df.groupby("Nombre"):
    lotes = group["Lote"].dropna().unique()
    sacrificios = group["F_Sacrificio"].dropna().unique()
    despieces = group["D_Prod_Desp"].dropna().unique()
    caducidades = group["F_Caducidad"].dropna().unique()

    result[nombre] = {
        "Lote(s)": ', '.join(lotes) + ',' if len(lotes) else '',
        "F.Sacrificio": ', '.join(sacrificios) + ',' if len(sacrificios) else '',
        "D.prod/desp": ', '.join(despieces) + ',' if len(despieces) else '',
        "F.caducidad": ', '.join(caducidades) + ',' if len(caducidades) else ''
    }

# Step 8: Print summary
for nombre, data in result.items():
    print(f"\nNombre: {nombre}")
    print(f"Lote(s): {data['Lote(s)']}")
    print(f"F.Sacrificio: {data['F.Sacrificio']}")
    print(f"D.prod/desp: {data['D.prod/desp']}")
    print(f"F.caducidad: {data['F.caducidad']}")

# Step 9: Convert result to DataFrame
df_result = pd.DataFrame.from_dict(result, orient='index').reset_index()
df_result = df_result.rename(columns={"index": "Nombre"})

# Step 10: Export summary to Excel and download
df_result.to_excel("resumen_por_nombre.xlsx", index=False)
files.download("resumen_por_nombre.xlsx")
