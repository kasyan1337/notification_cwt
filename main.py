import os
import subprocess
import sys
from tkinter import Tk, messagebox

# Define the paths to the scripts
base_dir = os.path.dirname(os.path.abspath(__file__))
script_dir = os.path.join(base_dir, 'src')
main_script = os.path.join(base_dir, 'main.py')
convert_script = os.path.join(script_dir, 'convert.py')
sort_script = os.path.join(script_dir, 'sort.py')
notify_script = os.path.join(script_dir, 'notify.py')

# GitHub repository URL and branch
GIT_REMOTE_URL = "git@github.com:kasyan1337/notification_cwt.git"
GIT_BRANCH = "develop"

# Ensure the Python executable is the same as the one running main.py
python_executable = sys.executable


# Function to check for script updates from GitHub
def update_scripts():
    try:
        # Navigate to the project folder and run 'git pull' to update the files
        print(f"Checking for updates from GitHub ({GIT_REMOTE_URL})...")
        result = subprocess.run(
            ["git", "pull", "origin", GIT_BRANCH],
            cwd=base_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if "Already up to date." in result.stdout.decode('utf-8'):
            print("No updates available. Scripts are already up to date.")
        else:
            print("Updates were fetched successfully. Restart the script to apply changes.")
            messagebox.showinfo("Update",
                                "New updates were fetched. Please restart the script to apply the latest changes.")
            sys.exit(0)  # Exit to allow for the restart with updated files
    except Exception as e:
        print(f"Failed to check for updates: {e}")
        messagebox.showerror("Error", f"Failed to check for updates: {e}")


# Function to run the scripts (for demonstration, runs only main script)
def run_scripts():
    try:
        subprocess.run([python_executable, main_script], check=True)
        subprocess.run([python_executable, convert_script], check=True)
        subprocess.run([python_executable, sort_script], check=True)
        subprocess.run([python_executable, notify_script], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running scripts: {e}")
        messagebox.showerror("Error", f"Error running scripts: {e}")


if __name__ == "__main__":
    root = Tk()
    root.withdraw()  # Hide the root window

    # Step 1: Check for updates from GitHub
    update_scripts()

    # Step 2: Run the main functionality
    run_scripts()