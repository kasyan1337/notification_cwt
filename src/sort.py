import os
from datetime import datetime, timedelta

import pandas as pd

# Define relative paths based on the script's location
base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))  # Root project directory
data_folder = os.path.join(base_dir, 'data')  # Path to the data folder
input_excel_file = os.path.join(data_folder, 'database_excel.xlsx')  # Input Excel file
expire_soon_file = os.path.join(data_folder, 'expire_soon.xlsx')  # Output file for expiring soon
already_expired_file = os.path.join(data_folder, 'already_expired.xlsx')  # Output file for already expired

if __name__ == '__main__':
    # Load the Excel file into a DataFrame
    # Load the Excel file into a DataFrame
    df = pd.read_excel(input_excel_file)

    # Define today's date and the date range for the next 3 months
    today = datetime.now()
    three_months_later = today + timedelta(days=90)

    # Filter the DataFrame
    filtered_df = df[pd.to_datetime(df['koniecplatnosti'], format='%m/%d/%y %H:%M:%S') <= three_months_later]

    # Make an explicit copy of the filtered DataFrame
    filtered_df = filtered_df.copy()

    # Calculate the number of days from now without triggering the warning
    filtered_df['Number_of_days_from_now'] = filtered_df['koniecplatnosti'].apply(
        lambda x: (pd.to_datetime(x) - today).days)

    # Split into two DataFrames: one for expired (negative days), one for expiring soon (zero or positive days)
    expire_soon_df = filtered_df[filtered_df['Number_of_days_from_now'] >= 0]
    already_expired_df = filtered_df[filtered_df['Number_of_days_from_now'] < 0]

    # Create a new DataFrame with the required columns and formatting
    output_data_soon = {
        'Number': range(1, len(expire_soon_df) + 1),  # 1, 2, 3, etc.
        'Identification': expire_soon_df['Identifikácia'],  # Column A
        'Index': expire_soon_df['index'],  # Column B
        'Name': expire_soon_df['meno'],  # Column E
        'Surname': expire_soon_df['priezvisko'],  # Column F
        'Method': expire_soon_df['metoda'],  # Column G
        'Level': expire_soon_df['stupeň'],  # Column H (Slovak diacritic)
        'Expiry_date': expire_soon_df['koniecplatnosti'].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%y')),
        'Number_of_days_from_now': expire_soon_df['Number_of_days_from_now'],
        'Training_center': expire_soon_df['školiace stredisko'],  # Column AC
    }

    # Data for already expired
    output_data_expired = {
        'Number': range(1, len(already_expired_df) + 1),  # 1, 2, 3, etc.
        'Identification': already_expired_df['Identifikácia'],  # Column A
        'Index': already_expired_df['index'],  # Column B
        'Name': already_expired_df['meno'],  # Column E
        'Surname': already_expired_df['priezvisko'],  # Column F
        'Method': already_expired_df['metoda'],  # Column G
        'Level': already_expired_df['stupeň'],  # Column H (Slovak diacritic)
        'Expiry_date': already_expired_df['koniecplatnosti'].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%y')),
        'Number_of_days_from_now': already_expired_df['Number_of_days_from_now'],
        'Training_center': already_expired_df['školiace stredisko'],  # Column AC
    }

    # Convert the dictionaries to DataFrames
    expire_soon_output_df = pd.DataFrame(output_data_soon)
    already_expired_output_df = pd.DataFrame(output_data_expired)

    # Sort both DataFrames by 'Number_of_days_from_now'
    expire_soon_output_df = expire_soon_output_df.sort_values(by='Number_of_days_from_now', ascending=True)
    already_expired_output_df = already_expired_output_df.sort_values(by='Number_of_days_from_now', ascending=False)

    # Re-assign the 'Number' column to ensure sequential numbering after sorting
    expire_soon_output_df['Number'] = range(1, len(expire_soon_output_df) + 1)
    already_expired_output_df['Number'] = range(1, len(already_expired_output_df) + 1)

    # Write the soon-to-expire DataFrame to an Excel file
    if not expire_soon_output_df.empty:
        with pd.ExcelWriter(expire_soon_file, engine='openpyxl') as writer:
            expire_soon_output_df.to_excel(writer, index=False, sheet_name='Expire Soon')

    # Write the already expired DataFrame to an Excel file
    if not already_expired_output_df.empty:
        with pd.ExcelWriter(already_expired_file, engine='openpyxl') as writer:
            already_expired_output_df.to_excel(writer, index=False, sheet_name='Already Expired')

    print(f"Expire Soon data saved to {expire_soon_file}")
    print(f"Already Expired data saved to {already_expired_file}")
