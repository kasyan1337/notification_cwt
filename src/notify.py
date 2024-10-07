import os
from tkinter import Tk, Checkbutton, Button, BooleanVar, messagebox, Canvas, Scrollbar, Frame, Label
from tkinter.font import Font, ITALIC
import pandas as pd
from datetime import datetime

# Define relative paths based on the script's location
base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))  # Root project directory
data_folder = os.path.join(base_dir, 'data')  # Path to the data folder
expire_soon_file = os.path.join(data_folder, 'expire_soon.xlsx')  # Expiring soon file
do_not_notify_file = os.path.join(data_folder, 'do_not_notify.xlsx')  # Do not notify file

# Load the Excel files into DataFrames
expire_soon_df = pd.read_excel(expire_soon_file)

# If the do_not_notify file exists, load it, otherwise create an empty DataFrame
if os.path.exists(do_not_notify_file):
    do_not_notify_df = pd.read_excel(do_not_notify_file)
else:
    do_not_notify_df = pd.DataFrame(columns=expire_soon_df.columns)


# Function to check if a candidate is already in the do_not_notify file
def is_in_do_not_notify(candidate_id):
    return not do_not_notify_df[do_not_notify_df['Index'] == candidate_id].empty


# Function to add selected candidates to the do_not_notify file and update expire_soon.xlsx
def add_to_do_not_notify(selected_candidates):
    global do_not_notify_df, expire_soon_df
    # Add the selected candidates to the do_not_notify DataFrame
    for candidate_id in selected_candidates:
        # Use .loc[] to ensure correct DataFrame modifications
        candidate_row = expire_soon_df.loc[expire_soon_df['Index'] == candidate_id].copy()

        # Set 'Do_not_notify' column to 1 before adding to do_not_notify_df
        candidate_row['Do_not_notify'] = 1
        do_not_notify_df = pd.concat([do_not_notify_df, candidate_row], ignore_index=True)

        # Set 'Do_not_notify' column to 1 for this candidate in expire_soon_df
        expire_soon_df.loc[expire_soon_df['Index'] == candidate_id, 'Do_not_notify'] = 1

    # Save the updated DataFrame to the do_not_notify file
    with pd.ExcelWriter(do_not_notify_file, engine='openpyxl') as writer:
        do_not_notify_df.to_excel(writer, index=False)

    # Save the updated expire_soon_df back to expire_soon.xlsx
    with pd.ExcelWriter(expire_soon_file, engine='openpyxl') as writer:
        expire_soon_df.to_excel(writer, index=False)

    messagebox.showinfo("Success", "Selected candidates added to the Do Not Notify list.")


# Function to create a styled message based on Number_of_days_from_now
def create_message(row):
    days = row['Number_of_days_from_now']
    expiry_date = row['Expiry_date']

    if days == 0:
        return f"vyprší dnes, dňa {expiry_date}"
    elif days == 1:
        return f"vyprší o {days} deň, dňa {expiry_date}"
    elif days in [2, 3, 4]:
        return f"vyprší o {days} dni, dňa {expiry_date}"
    else:
        return f"vyprší o {days} dní, dňa {expiry_date}"


# Function to check if today is in the second half of December
def check_database_update_notification():
    today = datetime.now()
    if today.month == 12 and today.day >= 1:
        return True
    return False


# Create the pop-up window
def create_notification_window():
    root = Tk()
    root.title("C-WT Expiring Soon Notification")

    # Set the window to be resizable
    root.resizable(True, True)  # Allow window resizing in both directions

    # Set a larger initial window size
    root.geometry("1000x700")  # Increase width and height

    # Set up a scrollable canvas
    canvas = Canvas(root)
    scrollbar = Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas)

    # Configure the scrollable area
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")
        )
    )

    # Create a window inside the canvas
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=1366)  # Adjust the width to fit content

    # Configure canvas scrolling
    canvas.configure(yscrollcommand=scrollbar.set)

    # Fix scrolling issue by binding to the canvas only
    def on_mouse_wheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    # Bind mousewheel to canvas for better scroll experience
    canvas.bind_all("<MouseWheel>", on_mouse_wheel)

    # Store the candidate checkboxes
    check_vars = []

    # Font for identifiers (bold text) and normal text
    bold_font = Font(family="Helvetica", size=12, weight="bold")
    normal_font = Font(family="Helvetica", size=12)

    # Font and color for the red parts (days and expiry date)
    warning_font = Font(family="Helvetica", size=12, weight="bold")
    warning_color = "red"
    mild_warning_color = "orange"

    # Create checkboxes and labels for each candidate that is not in the do_not_notify list
    for index, row in expire_soon_df.iterrows():
        candidate_id = row['Index']

        if not is_in_do_not_notify(candidate_id):
            var = BooleanVar()
            check_vars.append((candidate_id, var))
            days_message = create_message(row)

            # Create a frame to hold the checkbox and the message
            row_frame = Frame(scrollable_frame)
            row_frame.pack(fill='x', pady=5)

            # Add checkbox before the message
            Checkbutton(row_frame, text="", variable=var).pack(side='left')

            # Display the message in parts, with styles applied only to certain elements
            Label(row_frame, text=f"{row['Number']}  -", font=bold_font).pack(side='left')
            Label(row_frame, text=f"{row['Identification']} Certifikát pod indexom ", font=normal_font).pack(
                side='left')
            Label(row_frame, text=f"{row['Index']}", font=bold_font).pack(side='left')
            Label(row_frame, text=f" uchádzača ", font=normal_font).pack(side='left')
            Label(row_frame, text=f"{row['Name']} {row['Surname']}", font=bold_font).pack(side='left')
            Label(row_frame, text=f" na metódu ", font=normal_font).pack(side='left')
            Label(row_frame, text=f"{row['Method']}", font=bold_font).pack(side='left')
            Label(row_frame, text=f" stupeň ", font=normal_font).pack(side='left')
            Label(row_frame, text=f"{row['Level']}", font=bold_font).pack(side='left')

            # Highlight Number_of_days_from_now and Expiry_date in red or orange
            days = row['Number_of_days_from_now']
            if int(days) <= 30:
                Label(row_frame, text=f" {days_message}", fg=warning_color, font=warning_font).pack(side='left')
            elif 31 <= int(days) <= 60:
                Label(row_frame, text=f" {days_message}", fg=mild_warning_color, font=warning_font).pack(side='left')
            else:
                Label(row_frame, text=f" {days_message}", font=warning_font).pack(side='left')

    # Function to handle the OK button click
    def on_ok():
        selected_candidates = [candidate_id for candidate_id, var in check_vars if var.get()]
        if selected_candidates:
            add_to_do_not_notify(selected_candidates)
        root.destroy()

    # Create a frame for buttons at the bottom
    button_frame = Frame(root)
    button_frame.pack(pady=10, side="bottom")

    # Add OK and Cancel buttons, centered at the bottom
    Button(button_frame, text="OK", command=on_ok, padx=20, pady=5).pack(side='left', padx=10)
    Button(button_frame, text="Cancel", command=root.destroy, padx=20, pady=5).pack(side='right', padx=10)

    # Add the italic message at the bottom
    italic_font = Font(family="Helvetica", size=10, slant=ITALIC)
    Label(root, text="The database root file needs to be updated every year.\n Date of the last update: 8.10.2024", font=italic_font).pack(pady=10)

    # Check if today is in the second half of December and show the update message
    if check_database_update_notification():
        Label(root, text="Please contact Kasim to update the database file by the end of December.", font=italic_font, fg="red").pack(pady=5)

    # Pack the canvas and scrollbar
    canvas.pack(side="left", fill="both", expand=True)  # Expand the canvas to fit the window
    scrollbar.pack(side="right", fill="y")

    root.mainloop()


if __name__ == '__main__':
    # Run the notification window
    create_notification_window()