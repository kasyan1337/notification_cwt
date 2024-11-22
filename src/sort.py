import os
from datetime import datetime, timedelta

import pandas as pd

# Define relative paths based on the script's location
base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))  # Root project directory
data_folder = os.path.join(base_dir, 'data')  # Path to the data folder
input_excel_file = os.path.join(data_folder, 'database_excel.xlsx')  # Input Excel file

# Output files for other training centers
expire_soon_file = os.path.join(data_folder, 'expire_soon.xlsx')  # Output file for expiring soon
already_expired_file = os.path.join(data_folder, 'already_expired.xlsx')  # Output file for already expired

# Output files for the 'ZLY' training center
expire_soon_file_zly = os.path.join(data_folder, 'expire_soon_zly.xlsx')  # Output file for 'ZLY' expiring soon
already_expired_file_zly = os.path.join(data_folder, 'already_expired_zly.xlsx')  # Output file for 'ZLY' already expired

if __name__ == '__main__':
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_file)

    # Define today's date and the date range for the next 3 months
    today = datetime.now()
    three_months_later = today + timedelta(days=90)

    # Ensure 'koniecplatnosti' is in datetime format
    df['koniecplatnosti'] = pd.to_datetime(df['koniecplatnosti'], errors='coerce')

    # Filter the DataFrame for entries within the next 3 months
    filtered_df = df[df['koniecplatnosti'] <= three_months_later].copy()

    # Calculate the number of days from now
    filtered_df['Number_of_days_from_now'] = filtered_df['koniecplatnosti'].apply(
        lambda x: (x - today).days)

    # Fill missing 'Training_center' values with empty strings
    filtered_df['Training_center'] = filtered_df['školiace stredisko'].fillna('')

    # Split into two DataFrames: one for expired (negative days), one for expiring soon (zero or positive days)
    expire_soon_df = filtered_df[filtered_df['Number_of_days_from_now'] >= 0].copy()
    already_expired_df = filtered_df[filtered_df['Number_of_days_from_now'] < 0].copy()

    # Split each DataFrame based on 'Training_center' being 'ZLY' or not
    def split_by_training_center(df):
        df['Training_center_upper'] = df['Training_center'].str.upper()
        df_zly = df[df['Training_center_upper'] == 'ZLY'].copy()
        df_rest = df[df['Training_center_upper'] != 'ZLY'].copy()
        return df_zly, df_rest

    expire_soon_df_zly, expire_soon_df_rest = split_by_training_center(expire_soon_df)
    already_expired_df_zly, already_expired_df_rest = split_by_training_center(already_expired_df)

    # Function to process and save DataFrame
    def process_and_save(df_to_process, output_file, sort_ascending):
        if df_to_process.empty:
            print(f"No data to save to {output_file}")
            return

        # Create a new DataFrame with the required columns and formatting
        output_data = {
            'Number': range(1, len(df_to_process) + 1),  # 1, 2, 3, etc.
            'Identification': df_to_process['Identifikácia'],  # Column A
            'Index': df_to_process['index'],  # Column B
            'Name': df_to_process['meno'],  # Column E
            'Surname': df_to_process['priezvisko'],  # Column F
            'Method': df_to_process['metoda'],  # Column G
            'Level': df_to_process['stupeň'],  # Column H (Slovak diacritic)
            'Expiry_date': df_to_process['koniecplatnosti'].dt.strftime('%d/%m/%y'),
            'Number_of_days_from_now': df_to_process['Number_of_days_from_now'],
            'Training_center': df_to_process['Training_center'],  # Column AC
        }

        output_df = pd.DataFrame(output_data)

        # Sort DataFrame by 'Number_of_days_from_now'
        output_df = output_df.sort_values(by='Number_of_days_from_now', ascending=sort_ascending)

        # Re-assign the 'Number' column to ensure sequential numbering after sorting
        output_df['Number'] = range(1, len(output_df) + 1)

        # Write the DataFrame to an Excel file
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False)

        print(f"Data saved to {output_file}")

    # Process and save DataFrames
    # For expire_soon_df_rest
    process_and_save(expire_soon_df_rest, expire_soon_file, sort_ascending=True)
    # For expire_soon_df_zly
    process_and_save(expire_soon_df_zly, expire_soon_file_zly, sort_ascending=True)
    # For already_expired_df_rest
    process_and_save(already_expired_df_rest, already_expired_file, sort_ascending=False)
    # For already_expired_df_zly
    process_and_save(already_expired_df_zly, already_expired_file_zly, sort_ascending=False)