Here’s a sample README.md file for your project:

Expiring Soon Notification System

This project is a notification system that alerts users about certificates that are about to expire. The system reads data from Excel files, displays a list of certificates that will expire soon, and allows users to select certain candidates to exclude from future notifications. It also includes a reminder to update the database root file every year in the second half of December.

Features

	•	Expiring Soon Notifications: The system reads data from an Excel file (expire_soon.xlsx) and displays certificates that are about to expire.
	•	Do Not Notify List: Users can select certificates to be added to a “Do Not Notify” list. These entries are stored in the do_not_notify.xlsx file.
	•	Database Update Reminder: If today’s date is in the second half of December, the system reminds users to update the database root file for the upcoming year.
	•	Scrollable GUI: The system features a scrollable window with checkboxes to select candidates, making it user-friendly even with a large list of certificates.
	•	Customizable UI: The user interface includes styled messages and color-coded warnings for certificates expiring soon.

Folder Structure

Notification_cwt/
├── data/
│   ├── expire_soon.xlsx       # Excel file containing expiring certificates
│   ├── do_not_notify.xlsx     # Excel file containing candidates marked as "Do Not Notify"
├── src/
│   ├── main.py                # Main script that runs the notification system
│   ├── notify.py              # Script handling notification logic
│   ├── sort.py                # Sorting logic for certificates
│   ├── convert.py             # Database conversion script
├── README.md                  # Documentation for the project

Prerequisites

	•	Python 3.6+
	•	Required Python packages:
	•	pandas
	•	openpyxl
	•	tkinter (usually comes pre-installed with Python)

Installation

	1.	Clone the repository:

git clone https://github.com/kasyan1337/notification_cwt.git
cd notification_cwt


	2.	Install the required Python libraries:

pip install pandas openpyxl


	3.	Ensure that you have the necessary Excel files (expire_soon.xlsx, do_not_notify.xlsx) in the data folder.

How to Use

	1.	Run the System:
Run the main.py file to start the notification system:

python src/main.py


	2.	Notification Window:
	•	The system will open a pop-up window displaying all certificates that will expire soon.
	•	You can scroll through the list and select certificates you don’t want to be notified about in the future by checking the corresponding checkboxes.
	•	Click “OK” to update the do_not_notify.xlsx file with your selections.
	3.	Reminder to Update the Database:
	•	If the current month December, the system will notify users that the database root file needs to be updated for the next year. This reminder is shown at the bottom of the window in red text.

File Details

1. expire_soon.xlsx

	•	This file contains a list of certificates that are expiring soon. It must include columns like:
	•	Index: Unique identifier for the certificate.
	•	Name, Surname: Candidate’s name.
	•	Method, Level: Certification details.
	•	Number_of_days_from_now: The number of days until the certificate expires.

2. do_not_notify.xlsx

	•	This file stores the list of candidates who should not receive notifications for upcoming expirations.
	•	When candidates are selected in the notification window, they are added to this file and will not appear in future notifications.

Customization

Database Update Notification

	•	The system will display a red warning message in December reminding users to update the database for the upcoming year. This message can be customized in the main.py file if required.

License

This project is licensed under the MIT License. See the LICENSE file for details.

Contact

If you encounter any issues or have any questions, feel free to contact me:

	•	Email: kjanci@c-wt.sk
	•	GitHub: kasyan1337

This README.md file provides an overview of your project, including instructions for setup, usage, and a description of how the system works. You can modify it as needed to suit your preferences.