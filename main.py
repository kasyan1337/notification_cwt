import os
import subprocess
import sys
import tkinter as tk
from tkinter import messagebox

# Define the paths to the scripts
base_dir = os.path.dirname(os.path.abspath(__file__))
script_dir = os.path.join(base_dir, 'src')
convert_script = os.path.join(script_dir, 'convert.py')
sort_script = os.path.join(script_dir, 'sort.py')
notify_script = os.path.join(script_dir, 'notify.py')

# Ensure the Python executable is the same as the one running main.py
python_executable = sys.executable


# Function to show an error message with contact info
def show_error_message():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    contact_info = ("An error occurred. Please contact me at:\n"
                    "Email: kjanci@c-wt.sk\n"
                    "WhatsApp: +421 944 118 730\n"
                    "GitHub: https://github.com/kasyan1337")
    messagebox.showerror("Error", contact_info)
    root.destroy()


# Run the scripts in order with error handling
try:
    subprocess.run([python_executable, convert_script], check=True)
    subprocess.run([python_executable, sort_script], check=True)
    subprocess.run([python_executable, notify_script], check=True)
except Exception as e:
    print(f"Error occurred: {e}")
    show_error_message()  # Show the pop-up window with contact info
