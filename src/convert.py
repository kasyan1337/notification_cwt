import subprocess
import pandas as pd
import os
import shutil
from io import StringIO  # Import StringIO from the io module

# Define relative paths based on the script's location (assuming it's in the "src" folder)
base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))  # Root project directory
data_folder = os.path.join(base_dir, 'data')  # Relative path to the data folder
database_root_folder = os.path.join(base_dir, 'database_root')  # Folder for the .mdb file
data_mdb_file = os.path.join(database_root_folder, 'cop.mdb')  # Path to copy the .mdb file in the database_root folder
excel_file = os.path.join(data_folder, 'database_excel.xlsx')  # Path for the Excel file

# Function to copy MDB file to database_root folder
def copy_mdb_to_data(source_mdb_file, data_mdb_file):
    try:
        # Check if database_root folder exists, if not create it
        if not os.path.exists(database_root_folder):
            os.makedirs(database_root_folder)

        # Copy the MDB file, overwrite if it already exists
        shutil.copyfile(source_mdb_file, data_mdb_file)
        print(f"Copied {source_mdb_file} to {data_mdb_file}")
    except Exception as e:
        print(f"Error copying MDB file: {e}")

# Function to export MDB table to Excel (no CSV)
def convert_mdb_to_excel(mdb_file, excel_file, table_name):
    try:
        # Run mdb-export to export a table from .mdb to CSV format in memory
        result = subprocess.run(['mdb-export', mdb_file, table_name], capture_output=True, text=True)

        # Convert the CSV content to DataFrame using StringIO
        csv_data = result.stdout
        df = pd.read_csv(StringIO(csv_data))

        # Save DataFrame to Excel, overwrite if it exists
        with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name="Certifikát")
        print(f"Data converted to {excel_file} with sheet name 'Certifikát'")
    except Exception as e:
        print(f"Error converting to Excel: {e}")

# Main logic to copy MDB file and then convert to Excel
try:
    # Step 1: Specify the source path of the .mdb file
    # Assuming the cop.mdb file is already in the project's folder `database_root`
    source_mdb_file = os.path.join(database_root_folder, 'cop.mdb')

    # Step 2: Copy the .mdb file to the database_root folder
    copy_mdb_to_data(source_mdb_file, data_mdb_file)

    # Step 3: List the tables in the .mdb file using mdb-tables
    result = subprocess.run(['mdb-tables', data_mdb_file], capture_output=True, text=True)
    tables = result.stdout.split()

    # Check if any tables were found
    if tables:
        print("Tables found:", tables)
        # Export the table named "Certifikát"
        if "Certifikát" in tables:
            print(f'Exporting table "Certifikát" to Excel.')
            # Convert MDB to Excel, no CSV
            convert_mdb_to_excel(data_mdb_file, excel_file, "Certifikát")
        else:
            print('"Certifikát" table not found.')
    else:
        print("No tables found in the MDB file.")
except Exception as e:
    print(f"Error during operation: {e}")