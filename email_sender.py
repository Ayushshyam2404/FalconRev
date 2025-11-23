import os
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
from datetime import datetime

# Import modules
from email_listener import download_attachments
from data_processor import process_and_get_html_and_file


def send_daily_report():
    print("--- Starting Daily Automation Routine (Shadow Archive Strategy) ---")

    # 1. Check Inbox for TODAY'S new raw file
    print("\n>>> STEP 1: Checking Inbox for new input...")
    download_attachments()

    # 2. Generate Report
    print("\n>>> STEP 2: Generating Report...")
    result = process_and_get_html_and_file()

    if not result:
        print("No report generated. Aborting.")
        return

    html_content, csv_filepath = result
    report_date_id = datetime.now().strftime('%Y-%m-%d')

    # Load Config
    load_dotenv()
    EMAIL_USER = os.getenv('EMAIL_USER')
    EMAIL_PASS = os.getenv('EMAIL_PASS')
    EMAIL_TO = os.getenv('EMAIL_TO', EMAIL_USER)

    # Connect to Server once
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_USER, EMAIL_PASS)

        # --- EMAIL 1: THE BOSS'S REPORT (Clean, No Attachment) ---
        print(f"\n>>> STEP 3a: Sending Clean Report to {EMAIL_TO}...")
        msg_report = MIMEMultipart()
        msg_report['From'] = f"Orange Falcon Bot <{EMAIL_USER}>"
        msg_report['To'] = EMAIL_TO
        msg_report['Subject'] = f"Daily Sales Report - {report_date_id}"
        msg_report.attach(MIMEText(html_content, 'html'))

        server.send_message(msg_report)
        print("SUCCESS: Clean report sent.")

        # --- EMAIL 2: THE SHADOW ARCHIVE (To You, With CSV) ---
        # This ensures the script can find the data tomorrow!
        print(f"\n>>> STEP 3b: Archiving Data to Sent Items...")
        msg_archive = MIMEMultipart()
        msg_archive['From'] = f"Orange Falcon Bot <{EMAIL_USER}>"
        msg_archive['To'] = EMAIL_USER  # Send to self/bot
        msg_archive['Subject'] = f"DATA ARCHIVE - {report_date_id}"  # Special Subject Line
        msg_archive.attach(MIMEText("Archiving raw data for historical pickup calculation.", 'plain'))

        # Attach the CSV to this hidden email
        if csv_filepath and os.path.exists(csv_filepath):
            with open(csv_filepath, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(csv_filepath)}",
            )
            msg_archive.attach(part)

            server.send_message(msg_archive)
            print("SUCCESS: Data archived in Sent Items.")
        else:
            print("WARNING: No CSV file found to archive.")

        server.quit()
        print("\nAll Done! Boss got the report, You got the data.")

    except Exception as e:
        print(f"ERROR: Could not send email. {e}")


if __name__ == "__main__":
    send_daily_report()