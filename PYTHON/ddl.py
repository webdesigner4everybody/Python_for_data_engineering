import pandas as pd
import os

# Define file paths
input_file_path = r'C:\Users\SHEKSINH\Downloads\AT - Estrattori SVG da GDP V3.0.xlsx'  # Update this path
output_file_path = r'C:\Users\SHEKSINH\OneDrive - Capgemini\Documents'  # Update this path

# Check if input file exists
if not os.path.isfile(input_file_path):
    raise FileNotFoundError(f"The input file does not exist: {input_file_path}")

# Load Excel data into a dictionary of DataFrames, one for each sheet
excel_data = pd.read_excel(input_file_path, sheet_name=None)  # Load all sheets

# Initialize a dictionary to store SQL statements
sql_statements = {}

# Iterate through each sheet
for sheet_name, df in excel_data.items():
    # Clean up the DataFrame by dropping rows with all NaN values
    df = df.dropna(how='all')
    
    # Ensure the necessary columns are present
    required_columns = ['Table Name', 'Column Name', 'Data Type', 'Length/Size', 'Nullable', 'Default Value', 'Primary Key']
    for col in required_columns:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")

    # Iterate through the rows of the DataFrame
    for _, row in df.iterrows():
        table_name = row['Table Name']
        column_name = row['Column Name']
        data_type = row['Data Type']
        length = row['Length/Size']
        nullable = 'NOT NULL' if row['Nullable'].strip().lower() == 'no' else 'NULL'
        default = f"DEFAULT {row['Default Value']}" if pd.notna(row['Default Value']) else ''
        primary_key = 'PRIMARY KEY' if row['Primary Key'].strip().lower() == 'yes' else ''
        
        if table_name not in sql_statements:
            sql_statements[table_name] = []

        column_definition = f"{column_name} {data_type}"
        if length:
            column_definition += f"({length})"
        if default:
            column_definition += f" {default}"
        if primary_key:
            column_definition += f" {primary_key}"
        
        column_definition += f" {nullable}"
        
        sql_statements[table_name].append(column_definition)

# Generate and save the DDL scripts
with open(output_file_path, 'w') as file:
    for table, columns in sql_statements.items():
        file.write(f"CREATE TABLE {table} (\n")
        file.write(",\n".join(columns))
        file.write("\n);\n\n")

print(f"DDL scripts have been generated and saved to {output_file_path}")
