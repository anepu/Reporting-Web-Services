# Get-MessageTraceReport.ps1
# By Andrei Epure, Microsoft Ltd. 2025. Use at your own risk. No warranties are given.
# DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

"""
.SYNOPSIS Retrieves a message trace report from the Office 365 Reporting Web Service.

.DESCRIPTION This script demonstrates how to retrieve a message trace report from the Office 365 Reporting Web Service. Implements OAuth via auth code only (certificate auth not yet implemented, basic auth is not supported). Not all parameters available in the API are implemented, this script is purely for testing. 
For application registration instructions, please see https://learn.microsoft.com/en-us/previous-versions/office/developer/o365-enterprise-developers/jj984325(v=office.15)#register-your-application-in-azure-ad

Special thanks to David Barrett for the inspiration on this https://github.com/David-Barrett-MS/PowerShell/blob/main/Reporting%20Web%20Service/Get-MessageTraceReport.ps1
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import requests
from datetime import datetime, timedelta
import threading
import os
import subprocess
import sys

def get_message_trace_report():
    def background_task():
        try:
            # Gather user inputs
            app_id = app_id_entry.get().strip()
            tenant_id = tenant_id_entry.get().strip()
            app_secret = app_secret_entry.get().strip()
            start_date = start_date_entry.get_date().strftime('%Y-%m-%d')
            end_date = end_date_entry.get_date().strftime('%Y-%m-%d')
            save_path = save_path_var.get()

            # Validate inputs
            if not app_id or not tenant_id or not app_secret:
                messagebox.showerror("Error", "App ID, Tenant ID, and App Secret are required!")
                return

            if not start_date or not end_date:
                messagebox.showerror("Error", "Start Date and End Date are required!")
                return

            if not save_path:
                messagebox.showerror("Error", "Please select a folder to save the report!")
                return

            # Update the processing label
            processing_label.config(text="Processing... Please wait.")
            root.update()

            # Token acquisition
            auth_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
            token_data = {
                "grant_type": "client_credentials",
                "client_id": app_id,
                "client_secret": app_secret,
                "scope": "https://outlook.office365.com/.default",
            }
            response = requests.post(auth_url, data=token_data)
            response.raise_for_status()
            token = response.json().get("access_token")

            if not token:
                messagebox.showerror("Error", "Failed to retrieve OAuth token!")
                return

            # API call to Message Trace
            headers = {"Authorization": f"Bearer {token}"}
            base_url = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace"
            query_params = (
                f"$filter=StartDate eq datetime'{start_date}T00:00:00Z' and "
                f"EndDate eq datetime'{end_date}T23:59:59Z'"
            )
            url = f"{base_url}?{query_params}"

            response = requests.get(url, headers=headers)

            # Handle API response
            if response.status_code != 200:
                messagebox.showerror("Error", f"API call failed: {response.text}")
                return

            # Save the report
            output_path = os.path.join(save_path, f"MessageTraceReport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xml")
            with open(output_path, "wb") as file:
                file.write(response.content)

            # Update the processing label and show success
            processing_label.config(text="Processing complete!")
            messagebox.showinfo("Success", f"Report saved to {output_path}")

            # Open the folder containing the saved report
            subprocess.run(f'explorer /select,"{os.path.abspath(output_path)}"')

        except requests.exceptions.RequestException as req_err:
            messagebox.showerror("Error", f"Request error: {req_err}")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        finally:
            # Reset progress label after processing
            processing_label.config(text="")

    # Start background task in a separate thread to avoid blocking the UI
    threading.Thread(target=background_task, daemon=True).start()


def browse_folder():
    folder = filedialog.askdirectory()
    if folder:
        save_path_var.set(folder)


# GUI setup
root = tk.Tk()
root.title("Office 365 Reporting Web Services - MessageTrace")

# Set the window icon (favicon)
if getattr(sys, 'frozen', False):
    icon_path = os.path.join(sys._MEIPASS, 'Logo_RWS.ico')
else:
    icon_path = os.path.abspath("Logo_RWS.ico")

try:
    root.iconbitmap(icon_path)
except Exception as e:
    print(f"Error loading icon: {e}")

main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=0, column=0, sticky="NSEW")

# App ID
ttk.Label(main_frame, text="App ID:").grid(row=0, column=0, sticky="W")
app_id_entry = ttk.Entry(main_frame, width=50)
app_id_entry.grid(row=0, column=1, padx=5, pady=5)

# Tenant ID
ttk.Label(main_frame, text="Tenant ID:").grid(row=1, column=0, sticky="W")
tenant_id_entry = ttk.Entry(main_frame, width=50)
tenant_id_entry.grid(row=1, column=1, padx=5, pady=5)

# App Secret
ttk.Label(main_frame, text="App Secret:").grid(row=2, column=0, sticky="W")
app_secret_entry = ttk.Entry(main_frame, show="*", width=50)
app_secret_entry.grid(row=2, column=1, padx=5, pady=5)

# Start Date (using DateEntry)
ttk.Label(main_frame, text="Start Date:").grid(row=3, column=0, sticky="W")

# Calculate the allowed date range for Start Date
min_date = datetime.today().date() - timedelta(days=10)  # Maximum 10 days in the past
max_date = datetime.today().date()  # Up to today

# Set the StartDate calendar
start_date_entry = DateEntry(
    main_frame, 
    width=50, 
    date_pattern="yyyy-mm-dd", 
    mindate=min_date, 
    maxdate=max_date
)
start_date_entry.grid(row=3, column=1, padx=5, pady=5)

# End Date (using DateEntry)
ttk.Label(main_frame, text="End Date:").grid(row=4, column=0, sticky="W")
end_date_entry = DateEntry(main_frame, width=50, date_pattern="yyyy-mm-dd")
end_date_entry.grid(row=4, column=1, padx=5, pady=5)

# Default Save Path
default_save_path = os.getenv("TEMP")  # Use the system's temp folder
save_path_var = tk.StringVar(value=default_save_path)

# Save Path
ttk.Label(main_frame, text="Save Path:").grid(row=5, column=0, sticky="W")
save_path_entry = ttk.Entry(main_frame, textvariable=save_path_var, width=50)
save_path_entry.grid(row=5, column=1, padx=5, pady=5)
ttk.Button(main_frame, text="Browse", command=browse_folder).grid(row=5, column=2, padx=5, pady=5)

# Processing Label (initially empty)
processing_label = ttk.Label(main_frame, text="", foreground="red")
processing_label.grid(row=6, column=0, columnspan=3, pady=10, sticky="W")

# Generate Report Button
ttk.Button(main_frame, text="Generate Report", command=get_message_trace_report).grid(
    row=7, column=1, pady=10, sticky="E"
)

root.mainloop()
