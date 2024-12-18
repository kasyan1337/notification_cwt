import os
import ssl
import subprocess
import sys
import urllib.request
from datetime import datetime
from tkinter import Tk, messagebox

# Auto-update configuration: Set to 1 to enable auto-updates, 0 to disable
auto_update = 1  # Change to 0 to disable auto-update

# Define the paths to the scripts
base_dir = os.path.dirname(os.path.abspath(__file__))
script_dir = os.path.join(base_dir, 'src')
convert_script = os.path.join(script_dir, 'convert.py')
sort_script = os.path.join(script_dir, 'sort.py')
notify_script = os.path.join(script_dir, 'notify.py')
notify_zly_script = os.path.join(script_dir, 'notify_zly.py')

# GitHub raw URLs for the scripts (change branch if necessary)
RAW_BASE_URL = "https://raw.githubusercontent.com/kasyan1337/notification_cwt/main"
RAW_FILES = {
    "main.py": f"{RAW_BASE_URL}/main.py",
    "convert.py": f"{RAW_BASE_URL}/src/convert.py",
    "sort.py": f"{RAW_BASE_URL}/src/sort.py",
    "notify.py": f"{RAW_BASE_URL}/src/notify.py",
    "notify_zly.py": f"{RAW_BASE_URL}/src/notify_zly.py",
}

# Ensure the Python executable is the same as the one running main.py
python_executable = sys.executable

# Create an SSL context that bypasses SSL verification (not recommended for production environments)
ssl_context = ssl._create_unverified_context()


# Function to download and update the scripts from GitHub raw URLs
def update_scripts():
    try:
        updated = False
        for file_name, raw_url in RAW_FILES.items():
            local_file_path = os.path.join(base_dir, file_name) if file_name == 'main.py' else os.path.join(script_dir, file_name)
            print(f"Checking for updates for {file_name}...")

            # Download the latest version from GitHub raw
            with urllib.request.urlopen(raw_url, context=ssl_context) as response:
                latest_code = response.read().decode('utf-8')

            # Read the local file with explicit encoding
            if os.path.exists(local_file_path):
                with open(local_file_path, 'r', encoding='utf-8') as local_file:
                    local_code = local_file.read()
            else:
                local_code = ''

            if local_code != latest_code:
                # Update the file if it's different
                with open(local_file_path, 'w', encoding='utf-8') as local_file:
                    local_file.write(latest_code)
                print(f"{file_name} was updated.")
                updated = True
            else:
                print(f"{file_name} is already up to date.")

        if updated:
            messagebox.showinfo("Update", "Scripts were updated. Please restart the application to apply the changes.")
            sys.exit(0)  # Exit so the user can restart with the updated scripts
        else:
            print("All scripts are up to date.")
            return False  # No updates, proceed with the script execution
    except Exception as e:
        print(f"Failed to check for updates: {e}")
        messagebox.showerror("Error", f"Failed to check for updates: {e}")
        return False  # Continue if there was an error


# Function to run the scripts
def run_scripts():
    try:
        subprocess.run([python_executable, convert_script], check=True)
        subprocess.run([python_executable, sort_script], check=True)
        subprocess.run([python_executable, notify_script], check=True)

        # Check if today is Friday (weekday() == 4)
        today = datetime.now()
        if today.weekday() == 0:  # Monday is 0, Friday is 4
            print("Today is Monday. Running notify_zly.py...")
            subprocess.run([python_executable, notify_zly_script], check=True)
        else:
            print("Today is not Monday. Skipping notify_zly.py.")

    except subprocess.CalledProcessError as e:
        print(f"Error running scripts: {e}")
        messagebox.showerror("Error", f"Error running scripts: {e}")


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Hide the root window

    # Step 1: Check for updates from GitHub raw (only if auto_update is enabled)
    if auto_update == 1:
        updates_fetched = update_scripts()
    else:
        print("Auto-update is disabled.")
        updates_fetched = False  # Proceed to run scripts

    # Step 2: Run the main functionality if no updates were fetched
    if not updates_fetched:
        run_scripts()