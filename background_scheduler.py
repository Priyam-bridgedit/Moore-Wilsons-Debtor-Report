import schedule
import time
from MW import generate_report, send_report, generate_report_2, generate_report_3, generate_report_3_auto, generate_report_4_auto
from configparser import ConfigParser
from datetime import datetime, timedelta
import base64

config = ConfigParser()
config.read('config.ini')

# Decrypt the necessary fields
smtp_password = base64.b64decode(config.get('SMTP', 'password').encode()).decode()
time_to_send = config.get('SMTP', 'time')

def send_email():
    print("Running send_email...")

    # Current time
    now = datetime.now()

    # Compute the last Monday
    start_date_time = now - timedelta(days=now.weekday(), weeks=1)

    # Compute the last Sunday
    end_date_time = start_date_time + timedelta(days=6)

    start_date_time_str = start_date_time.strftime('%Y-%m-%d %H:%M:%S')
    end_date_time_str = end_date_time.strftime('%Y-%m-%d %H:%M:%S')

    # Generate the first report and get the DataFrame
    df1 = generate_report(start_date_time_str, "00:00:00", end_date_time_str, "23:59:59", save_to_file=False)

    # Generate the second report and get the DataFrame
    df2 = generate_report_2(start_date_time_str, end_date_time_str, save_to_file=False)

    # Generate the third report with last week's date range and get the DataFrame
    df3 = generate_report_3_auto()

    # Generate the fourth report with last week's date range and get the DataFrame
    df4 = generate_report_4_auto()

    # Send all the reports by email
    send_report(df1, df2, df3, df4, start_date_time_str, end_date_time_str)

    print("Finished running send_email.")




# Schedule the report generation and email sending
schedule.every().day.at(time_to_send).do(send_email)

# Keep the script running
while True:
    schedule.run_pending()
    time.sleep(1)
