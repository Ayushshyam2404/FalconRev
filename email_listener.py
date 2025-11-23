import os
from datetime import date, timedelta
from dotenv import load_dotenv
from imap_tools import MailBox, AND


def download_attachments():
    """
    Connects to inbox and syncs reports from the last 7 days.
    If a report from yesterday is missing, this will find and download it.
    """
    load_dotenv()
    EMAIL_USER = os.getenv('EMAIL_USER')
    EMAIL_PASS = os.getenv('EMAIL_PASS')
    EMAIL_IMAP_SERVER = os.getenv('EMAIL_IMAP_SERVER')

    download_folder = "csv_downloads"
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)

    print("Connecting to mailbox for sync...")

    try:
        with MailBox(EMAIL_IMAP_SERVER).login(EMAIL_USER, EMAIL_PASS, 'INBOX') as mailbox:

            # CALCULATE DATE WINDOW
            # We look back 7 days to ensure we have history for the "Pickup" report
            days_back = 7
            date_limit = date.today() - timedelta(days=days_back)

            print(f"Syncing reports since: {date_limit}...")

            # SEARCH: Fetch all emails (Read OR Unread) from the last 7 days
            # We filter by date, not by 'seen' status
            criteria = AND(date_gte=date_limit)
            messages = mailbox.fetch(criteria)

            download_count = 0

            for msg in messages:
                for att in msg.attachments:
                    if att.filename.lower().endswith(('.csv', '.txt')):

                        # CRITICAL: Check if we already have this file
                        filepath = os.path.join(download_folder, att.filename)

                        if not os.path.exists(filepath):
                            # We don't have it! Download it (Backfill)
                            with open(filepath, 'wb') as f:
                                f.write(att.payload)
                            print(f"  [NEW] Downloaded missing report: {att.filename}")
                            download_count += 1
                        else:
                            # We already have it, skip silently
                            pass

            if download_count == 0:
                print("All recent reports are already up to date.")
            else:
                print(f"Sync complete. Downloaded {download_count} new/missing files.")

    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    download_attachments()