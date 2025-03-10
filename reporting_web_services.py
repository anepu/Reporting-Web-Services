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

            # Report type selection
            report_type = report_type_combobox.get()

            if not app_id or not tenant_id or not app_secret:
                messagebox.showerror("Error", "App ID, Tenant ID, and App Secret are required!")
                return

            if not start_date or not end_date:
                messagebox.showerror("Error", "Start Date and End Date are required!")
                return

            if not save_path:
                messagebox.showerror("Error", "Please select a folder to save the report!")
                return

            # Validate if Message Trace Detail is selected and required fields are filled
            if report_type == "Message Trace Detail":
                sender_address = sender_address_entry.get().strip()
                recipient_address = recipient_address_entry.get().strip()
                message_trace_id = message_trace_id_entry.get().strip()

                if not sender_address or not recipient_address or not message_trace_id:
                    messagebox.showerror("Error", "Sender Address, Recipient Address, and Message Trace ID are mandatory for Message Trace Detail!")
                    return

            # Check if the directory exists, if not create it
            if not os.path.exists(save_path):
                os.makedirs(save_path)

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

            # API URL and query parameters setup based on report type
            headers = {"Authorization": f"Bearer {token}"}
            base_url = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTrace"

            if report_type == "Message Trace Detail":
                base_url = "https://reports.office365.com/ecp/reportingwebservice/reporting.svc/MessageTraceDetail"
                query_params = (
                    f"$filter=MessageTraceId eq guid'{message_trace_id}' and "
                    f"RecipientAddress eq '{recipient_address}' and "
                    f"SenderAddress eq '{sender_address}' and "
                    f"StartDate eq datetime'{start_date}T00:00:00Z' and "
                    f"EndDate eq datetime'{end_date}T23:59:59Z'"
                )
            else:  # Default to Message Trace
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

# Get the current date and calculate the max date for both Start Date and End Date
current_date = datetime.now()
max_past_date_for_start = current_date - timedelta(days=10)  # 10 days ago for start date
max_past_date_for_end = current_date + timedelta(days=10)  # 10 days in the future for end date

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

# Start Date (using DateEntry with a max date of 10 days ago from the current date)
ttk.Label(main_frame, text="Start Date:").grid(row=3, column=0, sticky="W")
start_date_entry = DateEntry(main_frame, width=50, date_pattern="yyyy-mm-dd", mindate=max_past_date_for_start, maxdate=current_date)
start_date_entry.grid(row=3, column=1, padx=5, pady=5)

# End Date (using DateEntry with a max date of 10 days ago from the current date and any day in the future)
ttk.Label(main_frame, text="End Date:").grid(row=4, column=0, sticky="W")
end_date_entry = DateEntry(main_frame, width=50, date_pattern="yyyy-mm-dd", maxdate=max_past_date_for_end)
end_date_entry.grid(row=4, column=1, padx=5, pady=5)

# Report Type (Dropdown)
ttk.Label(main_frame, text="Report Type:").grid(row=5, column=0, sticky="W", padx=5, pady=5)
report_type_combobox = ttk.Combobox(main_frame, values=["Message Trace", "Message Trace Detail"], state="readonly", width=48)
report_type_combobox.grid(row=5, column=1, padx=5, pady=5)

# Sender Address (visible only for Message Trace Detail)
ttk.Label(main_frame, text="Sender Address: *").grid(row=6, column=0, sticky="W", padx=5, pady=5)
sender_address_entry = ttk.Entry(main_frame, width=50)
sender_address_entry.grid(row=6, column=1, padx=5, pady=5)
sender_address_entry.grid_forget()  # Initially hidden

# Recipient Address (visible only for Message Trace Detail)
ttk.Label(main_frame, text="Recipient Address: *").grid(row=7, column=0, sticky="W", padx=5, pady=5)
recipient_address_entry = ttk.Entry(main_frame, width=50)
recipient_address_entry.grid(row=7, column=1, padx=5, pady=5)
recipient_address_entry.grid_forget()  # Initially hidden

# Message Trace ID (visible only for Message Trace Detail)
ttk.Label(main_frame, text="Message Trace ID: *").grid(row=8, column=0, sticky="W", padx=5, pady=5)
message_trace_id_entry = ttk.Entry(main_frame, width=50)
message_trace_id_entry.grid(row=8, column=1, padx=5, pady=5)
message_trace_id_entry.grid_forget()  # Initially hidden

# Default Save Path (hardcoded)
save_path_var = tk.StringVar(value=r"C:\temp\ReportingWebServices-logs")

# Save Path
ttk.Label(main_frame, text="Save Path:").grid(row=9, column=0, sticky="W", padx=5, pady=5)
save_path_entry = ttk.Entry(main_frame, textvariable=save_path_var, width=50)
save_path_entry.grid(row=9, column=1, padx=5, pady=5)
ttk.Button(main_frame, text="Browse", command=browse_folder).grid(row=9, column=2, padx=5, pady=5)

# Processing Label (initially empty)
processing_label = ttk.Label(main_frame, text="", foreground="red")
processing_label.grid(row=10, column=0, columnspan=3, pady=10, sticky="W")

# Generate Report Button
ttk.Button(main_frame, text="Generate Report", command=get_message_trace_report).grid(row=11, column=0, columnspan=3, pady=10)

# Show/Hide fields based on report type selection
def on_report_type_select(event):
    if report_type_combobox.get() == "Message Trace Detail":
        sender_address_entry.grid(row=6, column=1, padx=5, pady=5)
        recipient_address_entry.grid(row=7, column=1, padx=5, pady=5)
        message_trace_id_entry.grid(row=8, column=1, padx=5, pady=5)
    else:
        sender_address_entry.grid_forget()
        recipient_address_entry.grid_forget()
        message_trace_id_entry.grid_forget()

# Bind the event for report type selection
report_type_combobox.bind("<<ComboboxSelected>>", on_report_type_select)

root.mainloop()
