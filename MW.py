import io
import threading
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import pyodbc
from tkinter import filedialog
from configparser import ConfigParser
from tkinter import ttk
from tkcalendar import DateEntry  # Make sure to install this library
import schedule
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from io import BytesIO
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import date, datetime, timedelta
import re
from datetime import datetime, timedelta
from queue import Queue
import time
import tkinter as tk
from tkinter import ttk, filedialog
from io import StringIO
import base64
import smtplib
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
import matplotlib.pyplot as plt

from tkinter import filedialog
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog


server_entry = None
database_entry = None
username_entry = None
password_entry = None
smtp_server_entry = None
smtp_username_entry = None
smtp_password_entry = None
smtp_from_entry = None
to_email_entry = None
time_entry = None

# Define config globally
config = ConfigParser()
# Function to save both SQL Server and SMTP details to config.ini file
def save_config(config_window):
    config["DATABASE"] = {
        "server": base64.b64encode(server_entry.get().encode()).decode(),
        "database": base64.b64encode(database_entry.get().encode()).decode(),
        "username": base64.b64encode(username_entry.get().encode()).decode(),
        "password": base64.b64encode(password_entry.get().encode()).decode(),
    }

    config["SMTP"] = {
        "server": base64.b64encode(smtp_server_entry.get().encode()).decode(),
        "username": base64.b64encode(smtp_username_entry.get().encode()).decode(),
        "password": base64.b64encode(smtp_password_entry.get().encode()).decode(),
        "from": base64.b64encode(smtp_from_entry.get().encode()).decode(),
        "to": base64.b64encode(to_email_entry.get().encode()).decode(),
        "time": time_entry.get(),
    }

    with open("config.ini", "w") as configfile:
        config.write(configfile)

    status_label.config(text="Configuration saved successfully!", fg="green")
    config_window.destroy()


# Function to schedule the report generation and email sending
time_format = re.compile("^([0-1]?[0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?$")


def schedule_report():
    print("Running schedulereport...")
    global config
    config.read("config.ini")
    time_to_send = config.get("SMTP", "time")
    if not time_format.match(time_to_send):
        status_queue.put(
            (
                "Invalid time format! Please enter time in HH:MM or HH:MM:SS format.",
                "red",
            )
        )
        return

    # Compute "yesterday"
    now = datetime.now()
    start_date_time = now - timedelta(days=1)
    end_date_time = start_date_time  # The start and end dates are the same for "yesterday"

    start_date_time_str = start_date_time.strftime("%Y-%m-%d 00:00:00")
    end_date_time_str = start_date_time.strftime("%Y-%m-%d 23:59:59")

    # Calculate start and end date strings
    start_date_str = start_date_time.strftime("%Y-%m-%d")
    end_date_str = start_date_time.strftime("%Y-%m-%d")
    start_time_str = "00:00:00"
    end_time_str = "23:59:59"

    schedule.every().day.at(time_to_send).do(
        lambda: send_report(
            generate_report(
                start_date_str,
                start_time_str,
                end_date_str,
                end_time_str,
                start_date_time_str,
                end_date_time_str,
                save_to_file=False,
            ),
            
            start_date_time_str,
            end_date_time_str
        )
    )

    print("Report Scheduled")



# Create a queue for status updates
status_queue = Queue()

# Define window as a global variable
window = None

def push_data_to_exonet(df, start_date, start_time, end_date, end_time, start_date_time_str, end_date_time_str):
    try:
        config = ConfigParser()
        config.read("config.ini")
        server = base64.b64decode(config.get("DATABASE", "server").encode()).decode()
        
        # Instead of reading the database from config.ini, set it directly here
        database = "ToInfinityAndBeyond"  # Change the database to EXONET_TEST
        print(f"Target Database: {database}")  # Print the target database

        username = base64.b64decode(config.get("DATABASE", "username").encode()).decode()
        password = base64.b64decode(config.get("DATABASE", "password").encode()).decode()

        if username and password:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        else:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes"
        
        # Print the connection string (without the password for security reasons)
        safe_connection_string = connection_string.replace(password, "*****")
        print(f"Connection String: {safe_connection_string}")

        connection = pyodbc.connect(connection_string)
        cursor = connection.cursor()

        start_date_time = pd.to_datetime(start_date + " " + start_time)
        end_date_time = pd.to_datetime(end_date + " " + end_time)
        start_date_time_str = start_date_time.strftime("%Y-%m-%d %H:%M:%S")
        end_date_time_str = end_date_time.strftime("%Y-%m-%d %H:%M:%S")

        # After fetching data into df
        for _, row in df.iterrows():
            trans_date = row['Invoice Date']
            acc_no = row['A/C Number']
            inv_no = row['Invoice Number']
            sub_total = row['Amount']
            tax_total = row['GST']
            branch = row['Branch']
            station = row['Station']
            trans_no = row['TransNo']


            # Print data that's about to be inserted
            print(f"Inserting data: Invoice Date: {trans_date}, A/C Number: {acc_no}, Invoice Number: {inv_no}, Amount: {sub_total}, GST: {tax_total}, Branch: {branch}, Station: {station}, TransNo: {trans_no}")

            # Call the stored procedure
            cursor = connection.cursor()
            try:
                cursor.execute(
                    """EXEC ToInfinityAndBeyond.dbo.sp_Insert_ExoDRTRANS 
                    @TransDate = ?, 
                    @AccNo = ?, 
                    @InvNo = ?, 
                    @SubTotal = ?, 
                    @TaxTotal = ?, 
                    @Branch = ?, 
                    @Station = ?, 
                    @TransNo = ?""", 
                    trans_date, acc_no, inv_no, sub_total, tax_total, branch, station, trans_no
                )
                connection.commit()
            except Exception as e:
                # Handle exception, like logging it or raising it
                print(f"Failed to insert record for Invoice Number: {inv_no}. Error: {e}")
            finally:
                cursor.close()

        connection.commit()
        connection.close()

    except Exception as e:
        print(f"Error: {str(e)}")
        connection.rollback()


# Function to generate the report
def generate_report(
    start_date,
    start_time,
    end_date,
    end_time,
    start_date_time_str,
    end_date_time_str,
    save_to_file=True,
):
    print(f"Value of save_to_file: {save_to_file}")
    try:
        config = ConfigParser()
        config.read("config.ini")
        server = base64.b64decode(config.get("DATABASE", "server").encode()).decode()
        database = base64.b64decode(
            config.get("DATABASE", "database").encode()
        ).decode()
        username = base64.b64decode(
            config.get("DATABASE", "username").encode()
        ).decode()
        password = base64.b64decode(
            config.get("DATABASE", "password").encode()
        ).decode()

        if username and password:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};UID={username};PWD={password}"
        else:
            connection_string = f"DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes"

        connection = pyodbc.connect(connection_string)

        start_date_time = pd.to_datetime(start_date + " " + start_time)
        end_date_time = pd.to_datetime(end_date + " " + end_time)

        start_date_time_str = start_date_time.strftime("%Y-%m-%d %H:%M:%S")
        end_date_time_str = end_date_time.strftime("%Y-%m-%d %H:%M:%S")

        # Modify the query to include the date range filter
        query = f"""
	SELECT 
    TH.[TransNo],
    TH.Branch,
    TH.Station,
    TH.Receipt AS "Invoice Number",
    TH.Logged AS "Invoice Date",
    C.CField_Integer AS "A/C Number",
    TP.[Value] AS ChargedToAccount,
    TotalBeforeTax AS Amount,
    TotalAfterTax - TotalBeforeTax AS GST,
    TotalAfterTax AS Gross
FROM 
    AKPOS.dbo.[TransHeaders] TH WITH (NOLOCK)
JOIN 
    AKPOS.dbo.[Customers] C ON C.[Code] = TH.[Customer] AND C.CustType = 'C'
JOIN 
    AKPOS.dbo.[TransPayments] TP ON TP.[TransNo] = TH.[TransNo]
    AND TP.[Branch] = TH.[Branch]
    AND TP.[Station] = TH.[Station]
LEFT JOIN 
    [ToInfinityAndBeyond].dbo.[ExonetDebtorLink] ExoLink ON ExoLink.[Receipt] COLLATE Database_Default = TH.[Receipt] COLLATE Database_Default
    AND ExoLink.[Logged] = TH.[Logged]
    AND ExoLink.[TransNo] = TH.[TransNo]
    AND ExoLink.[Branch] = TH.[Branch]
    AND ExoLink.[Station] = TH.[Station]
WHERE 
    ExoLink.Receipt IS NULL
    AND TH.[TransType] = 'A'
    AND TH.[TransStatus] = 'C'
    AND TP.[MediaID] = 4
    AND TH.Logged BETWEEN '{start_date_time_str}' AND '{end_date_time_str}'

        """
        df = pd.read_sql_query(query, connection)

        # df['Invoice Number'] = "'" + df['Invoice Number'] + "'"

        # Create a new DataFrame for the modified report
        modified_df = pd.DataFrame(columns=df.columns)
        # Push the data to Exonet
        push_data_to_exonet(df, start_date, start_time, end_date, end_time, start_date_time_str, end_date_time_str)
        if save_to_file:
            file_path = filedialog.asksaveasfilename(defaultextension=".csv")
            if file_path:
                df.to_csv(file_path, index=False)
                status_label.config(
                    text="Report generated and saved successfully!", fg="green"
                )
            else:
                status_label.config(text="Report generation cancelled.", fg="red")

        

        return df

        connection.close()

    except Exception as e:
        status_queue.put((f"Error: {str(e)}", "red"))


from pandas import ExcelWriter


def send_report(df1, start_date_time_str, end_date_time_str):
    print("Running scheduled_email...")
    try:
        # Check if df1, df3, and df4 are not None
        if df1 is None:
            raise ValueError("df1 is None")

        # Create StringIO objects and save the DataFrames to them
        csv_buffer1 = StringIO()
        
        df1.to_csv(csv_buffer1, index=False)
        

        

        # Read SMTP details from config.ini file
        config = ConfigParser()
        config.read("config.ini")
        smtp_server = base64.b64decode(config.get("SMTP", "server").encode()).decode()
        smtp_username = base64.b64decode(
            config.get("SMTP", "username").encode()
        ).decode()
        smtp_password = base64.b64decode(
            config.get("SMTP", "password").encode()
        ).decode()
        smtp_from = base64.b64decode(config.get("SMTP", "from").encode()).decode()
        to_email = (
            base64.b64decode(config.get("SMTP", "to").encode()).decode().split(",")
        )

        with smtplib.SMTP_SSL(smtp_server, 465) as server:
            server.login(smtp_username, smtp_password)
            for to_address in to_email:
                to_address = to_address.strip()

                # Initialize the MIMEMultipart instance inside the loop
                msg = MIMEMultipart()
                msg["From"] = smtp_from
                msg["To"] = to_address
                msg["Subject"] = f"Debtor Daily Reports for {datetime.strptime(start_date_time_str, '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')}"


                body = "Please find the daily reports attached."
                msg.attach(MIMEText(body, "plain"))

                part1 = MIMEBase("application", "octet-stream")
                part1.set_payload(csv_buffer1.getvalue())
                encoders.encode_base64(part1)
                part1.add_header(
                    "Content-Disposition",
                    f"attachment; filename= MW Debtor Report.csv",
                )

                msg.attach(part1)

                server.send_message(msg)
                print(f"Reports sent to {to_address}")

    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback

        print(traceback.format_exc())


def generate_both_reports(
    start_date_str,
    start_time_str,
    end_date_str,
    end_time_str,
    current_date_time_str,
    previous_date_time_str,
):
    start_date_time_str = f"{start_date_str} {start_time_str}"
    end_date_time_str = f"{end_date_str} {end_time_str}"

    generate_report(
        start_date_str,
        start_time_str,
        end_date_str,
        end_time_str,
        current_date_time_str,
        previous_date_time_str,
    )


def start_scheduler():
    while True:
        schedule.run_pending()
        time.sleep(1)


# Function to schedule the report and start the scheduler
def schedule_and_start():
    schedule_report()
    scheduler_thread = threading.Thread(target=start_scheduler)
    scheduler_thread.start()


if __name__ == "__main__":
    schedule_and_start()


def open_config_window():
    global server_entry, database_entry, username_entry, password_entry
    global smtp_server_entry, smtp_username_entry, smtp_password_entry
    global smtp_from_entry, to_email_entry, time_entry

    config_window = tk.Toplevel(window)
    config_window.title("Update Configuration")

    # All your Labels and Entries come here

    # Server Label and Entry
    server_label = tk.Label(config_window, text="Server:")
    server_label.pack()
    server_entry = tk.Entry(config_window)
    server_entry.pack()

    # Database Label and Entry
    database_label = tk.Label(config_window, text="Database:")
    database_label.pack()
    database_entry = tk.Entry(config_window)
    database_entry.pack()

    # Username Label and Entry
    username_label = tk.Label(
        config_window, text="Username (leave blank for Windows Authentication):"
    )
    username_label.pack()
    username_entry = tk.Entry(config_window)
    username_entry.pack()

    # Password Label and Entry
    password_label = tk.Label(
        config_window, text="Password (leave blank for Windows Authentication):"
    )
    password_label.pack()
    password_entry = tk.Entry(config_window, show="*")
    password_entry.pack()

    # SMTP Server Label and Entry
    smtp_server_label = tk.Label(config_window, text="SMTP Server:")
    smtp_server_label.pack()
    smtp_server_entry = tk.Entry(config_window)
    smtp_server_entry.pack()

    # SMTP Username Label and Entry
    smtp_username_label = tk.Label(config_window, text="SMTP Username:")
    smtp_username_label.pack()
    smtp_username_entry = tk.Entry(config_window)
    smtp_username_entry.pack()

    # SMTP Password Label and Entry
    smtp_password_label = tk.Label(config_window, text="SMTP Password:")
    smtp_password_label.pack()
    smtp_password_entry = tk.Entry(config_window, show="*")
    smtp_password_entry.pack()

    # 'From' Email Address Label and Entry
    smtp_from_label = tk.Label(config_window, text="'From' Email Address:")
    smtp_from_label.pack()
    smtp_from_entry = tk.Entry(config_window)
    smtp_from_entry.pack()

    # 'To' Email Address Label and Entry
    to_email_label = tk.Label(config_window, text="'To' Email Address:")
    to_email_label.pack()
    to_email_entry = tk.Entry(config_window)
    to_email_entry.pack()

    # Time to Send Report Label and Entry
    time_entry_label = tk.Label(config_window, text="Time to Send Report (HH:MM):")
    time_entry_label.pack()
    time_entry = tk.Entry(config_window)
    time_entry.pack()

    # Save SMTP Config Button
    # Pass the config_window to the save_config function
    save_config_button = tk.Button(
        config_window, text="Save Config", command=lambda: save_config(config_window)
    )
    save_config_button.pack()


window = tk.Tk()
window.title("Moore Wilsons Debtor Report")
window.geometry("400x350")
window.configure(bg="#f0f0f0")  # Set background color

# Header Label
header_label = tk.Label(window, text="Moore Wilsons Debtor Report", font=("Helvetica", 16, "bold"), bg="#f0f0f0")
header_label.pack(pady=10)

# Frame for Date and Time Inputs
input_frame = tk.Frame(window, bg="#f0f0f0")
input_frame.pack(padx=20, pady=10)

# Start Date Input
start_date_label = tk.Label(input_frame, text="Start Date:", bg="#f0f0f0")
start_date_label.grid(row=0, column=0, padx=5, pady=5)
start_date_entry = DateEntry(input_frame)
start_date_entry.grid(row=0, column=1, padx=5, pady=5)

# Start Time Input
start_time_label = tk.Label(input_frame, text="Start Time (HH:MM):", bg="#f0f0f0")
start_time_label.grid(row=1, column=0, padx=5, pady=5)
start_time_entry = ttk.Entry(input_frame)
start_time_entry.grid(row=1, column=1, padx=5, pady=5)

# End Date Input
end_date_label = tk.Label(input_frame, text="End Date:", bg="#f0f0f0")
end_date_label.grid(row=2, column=0, padx=5, pady=5)
end_date_entry = DateEntry(input_frame)
end_date_entry.grid(row=2, column=1, padx=5, pady=5)

# End Time Input
end_time_label = tk.Label(input_frame, text="End Time (HH:MM):", bg="#f0f0f0")
end_time_label.grid(row=3, column=0, padx=5, pady=5)
end_time_entry = ttk.Entry(input_frame)
end_time_entry.grid(row=3, column=1, padx=5, pady=5)

# Generate Report Button
generate_report_button = tk.Button(
    window,
    text="Generate Report",
    command=lambda: generate_both_reports(
        start_date_entry.get(),
        start_time_entry.get(),
        end_date_entry.get(),
        end_time_entry.get(),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        (datetime.now() - timedelta(seconds=1)).strftime("%Y-%m-%d %H:%M:%S"),
    ),
    bg="#007acc",  # Set button color
    fg="white",    # Set text color
)
generate_report_button.pack(pady=10)

# Update Configuration Button
update_config_button = tk.Button(
    window,
    text="Update Configuration",
    command=open_config_window,
    bg="#4caf50",  # Set button color
    fg="white",    # Set text color
)
update_config_button.pack(pady=5)

# Status Label
status_label = tk.Label(window, text="", bg="#f0f0f0")
status_label.pack(pady=10)


# Start the GUI event loop
window.mainloop()