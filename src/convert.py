import os
import shutil
import pandas as pd
import platform

# Define relative paths based on the script's location
base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
data_folder = os.path.join(base_dir, 'data')
database_root_folder = os.path.join(base_dir, 'database_root')
data_mdb_file = os.path.join(database_root_folder, 'cop.mdb')
excel_file = os.path.join(data_folder, 'database_excel.xlsx')

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

if __name__ == '__main__':
    try:
        # Step 1: Specify the source path of the .mdb file
        source_mdb_file = os.path.join(database_root_folder, 'cop.mdb')

        # Step 2: Copy the .mdb file to the database_root folder
        copy_mdb_to_data(source_mdb_file, data_mdb_file)

        # Check if operating system is Windows
        if platform.system() == 'Windows':
            # Use pyodbc to read the MDB file
            import pyodbc

            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                r'DBQ=' + data_mdb_file + ';'
            )
            conn = pyodbc.connect(conn_str)
            sql = 'SELECT * FROM [Certifikát]'
            df = pd.read_sql(sql, conn)
            conn.close()

            # Save DataFrame to Excel
            with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name="Certifikát")
            print(f"Data converted to {excel_file} with sheet name 'Certifikát'")
        else:
            # Use mdb-tools commands on non-Windows systems
            import subprocess
            from io import StringIO

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
                    result = subprocess.run(['mdb-export', data_mdb_file, "Certifikát"], capture_output=True, text=True)
                    csv_data = result.stdout
                    df = pd.read_csv(StringIO(csv_data))

                    # Save DataFrame to Excel
                    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
                        df.to_excel(writer, index=False, sheet_name="Certifikát")
                    print(f"Data converted to {excel_file} with sheet name 'Certifikát'")
                else:
                    print('"Certifikát" table not found.')
            else:
                print("No tables found in the MDB file.")
    except Exception as e:
        print(f"Error during operation: {e}")