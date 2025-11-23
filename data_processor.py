import os
import pandas as pd
import glob
import re
from datetime import datetime, timedelta
from dotenv import load_dotenv
from imap_tools import MailBox, AND


def get_sorted_files():
    """Finds all local CSV/TXT files and sorts them by newest first."""
    files = glob.glob('csv_downloads/*.csv') + glob.glob('csv_downloads/*.txt')
    files.sort(key=os.path.getctime, reverse=True)
    return files


def extract_date_from_filename(filename):
    """Parses YYYY-MM-DD from filename."""
    match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)
    if match:
        return datetime.strptime(match.group(1), '%Y-%m-%d')
    return None


def fetch_previous_from_sent_folder(target_date):
    """
    Searches 'Sent Mail' for the SHADOW ARCHIVE email to get historical data.
    Target Date Format: YYYY-MM-DD
    """
    print(f"   > Searching Sent Items for archive: DATA ARCHIVE - {target_date}...")
    load_dotenv()
    EMAIL_USER = os.getenv('EMAIL_USER')
    EMAIL_PASS = os.getenv('EMAIL_PASS')
    EMAIL_IMAP_SERVER = os.getenv('EMAIL_IMAP_SERVER')

    try:
        # Note: Gmail's sent folder is usually '[Gmail]/Sent Mail'
        with MailBox(EMAIL_IMAP_SERVER).login(EMAIL_USER, EMAIL_PASS, '[Gmail]/Sent Mail') as mailbox:
            # Look for the special subject line we created in email_sender.py
            subject_query = f"DATA ARCHIVE - {target_date}"
            criteria = AND(subject=subject_query)

            # Fetch only the 1 most recent matching email
            for msg in mailbox.fetch(criteria, limit=1):
                for att in msg.attachments:
                    if att.filename.lower().endswith('.csv'):
                        save_path = os.path.join("csv_downloads", att.filename)

                        # Ensure folder exists
                        if not os.path.exists("csv_downloads"):
                            os.makedirs("csv_downloads")

                        with open(save_path, 'wb') as f:
                            f.write(att.payload)
                        print(f"   > Restored historical file: {att.filename}")
                        return save_path
    except Exception as e:
        print(f"   > Warning: Could not fetch from Sent folder: {e}")

    return None


def clean_dataframe(df):
    """Cleans currency strings ($1,000) into numbers (1000.0)."""
    if 'Date' in df.columns:
        df['DateObj'] = pd.to_datetime(df['Date'], errors='coerce')

    for col in ['Total Revenue', 'ADR', 'Total Rooms Sold Allocated']:
        if col in df.columns:
            # Remove $ and ,
            df[col] = df[col].astype(str).str.replace(r'[$,]', '', regex=True)
            # Convert to number, turn errors into 0
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    return df


def process_and_get_html_and_file():
    """
    Orchestrates the report generation:
    1. Identify Current File
    2. Hunt for Previous File (Local -> Sent Folder)
    3. Calculate Pickup
    4. Generate HTML
    Returns: (html_string, filepath_of_current_csv)
    """
    files = get_sorted_files()
    if not files:
        print("No files found locally!")
        return None

    current_file = files[0]
    print(f"Current File: {current_file}")

    # Determine Dates
    report_date = extract_date_from_filename(current_file)
    if not report_date:
        report_date = datetime.now()

    # Calculate Yesterday
    prev_date = report_date - timedelta(days=1)
    prev_date_str = prev_date.strftime('%Y-%m-%d')

    # --- STRATEGY: FIND PREVIOUS FILE ---
    # 1. Try Local Search first
    previous_file = None
    for f in files:
        if prev_date_str in f and f != current_file:
            previous_file = f
            break

    # 2. If not local, Try Sent Folder (Archive Search)
    if not previous_file:
        previous_file = fetch_previous_from_sent_folder(prev_date_str)

    try:
        # Load Current Data
        df_curr = clean_dataframe(pd.read_csv(current_file))

        # Load Previous Data (If found)
        if previous_file:
            print(f"Comparing against: {previous_file}")
            df_prev = clean_dataframe(pd.read_csv(previous_file))

            # Merge tables based on DateObj
            merged = pd.merge(
                df_curr,
                df_prev[['DateObj', 'Total Rooms Sold Allocated']],
                on='DateObj', how='left', suffixes=('', '_Prev')
            )
            # Calculate Pickup
            merged['Pickup'] = merged['Total Rooms Sold Allocated'] - merged['Total Rooms Sold Allocated_Prev'].fillna(
                0)
        else:
            print("No previous history found. Pickup is 0.")
            merged = df_curr.copy()
            merged['Pickup'] = 0

        # Filter for the next 7 Days
        start_date = report_date
        end_date = report_date + timedelta(days=6)
        final_df = merged[(merged['DateObj'] >= start_date) & (merged['DateObj'] <= end_date)].copy()
        final_df['Day'] = final_df['DateObj'].dt.strftime('%a')

        # --- HTML GENERATION (Orange Falcon Branding) ---
        COLOR_ORANGE = "#F05A28"
        COLOR_HEADER = "#404040"

        html = f"""
        <div style="font-family: Arial, sans-serif; color: #333;">
            <div style="border-bottom: 4px solid {COLOR_ORANGE}; padding-bottom: 10px; margin-bottom: 15px;">
                <h2 style="margin:0; color: #2D2D2D;">7-Day Pickup Report</h2>
                <span style="color:#888; font-size:12px;">Date: {report_date.strftime('%b %d, %Y')}</span>
            </div>
            <table style="width:100%; border-collapse:collapse; font-size:13px;">
                <thead style="background-color:{COLOR_HEADER}; color:white;">
                    <tr>
                        <th style="padding:8px;">Date</th>
                        <th style="padding:8px;">Day</th>
                        <th style="padding:8px;">Rooms</th>
                        <th style="padding:8px;">Rev</th>
                        <th style="padding:8px;">ADR</th>
                        <th style="padding:8px; background-color:{COLOR_ORANGE};">Pickup</th>
                    </tr>
                </thead>
                <tbody>
        """

        total_rooms = 0
        total_rev = 0
        total_pickup = 0

        for _, row in final_df.iterrows():
            total_rooms += row['Total Rooms Sold Allocated']
            total_rev += row['Total Revenue']
            total_pickup += row['Pickup']

            # Format Pickup logic
            pk = int(row['Pickup'])
            pk_style = "font-weight:bold;"
            pk_style += "color:green;" if pk > 0 else "color:red;" if pk < 0 else "color:#ccc;"

            html += f"""
                <tr style="border-bottom:1px solid #ddd;">
                    <td style="padding:8px; text-align:center;">{row['DateObj'].strftime('%m/%d')}</td>
                    <td style="padding:8px; text-align:center;">{row['Day']}</td>
                    <td style="padding:8px; text-align:right;">{int(row['Total Rooms Sold Allocated'])}</td>
                    <td style="padding:8px; text-align:right;">${row['Total Revenue']:,.0f}</td>
                    <td style="padding:8px; text-align:right;">${row['ADR']:,.2f}</td>
                    <td style="padding:8px; text-align:right; {pk_style}">{pk:+d}</td>
                </tr>
            """

        # Totals Calculation
        avg_adr = (total_rev / total_rooms) if total_rooms > 0 else 0

        html += f"""
                <tr style="background-color:#f4f4f4; font-weight:bold;">
                    <td colspan="2" style="padding:8px;">TOTAL</td>
                    <td style="padding:8px; text-align:right;">{int(total_rooms)}</td>
                    <td style="padding:8px; text-align:right;">${total_rev:,.0f}</td>
                    <td style="padding:8px; text-align:right;">${avg_adr:,.2f}</td>
                    <td style="padding:8px; text-align:right; color:{COLOR_ORANGE};">{int(total_pickup):+d}</td>
                </tr>
                </tbody>
            </table>
        </div>
        """

        return html, current_file

    except Exception as e:
        print(f"Error building report: {e}")
        import traceback
        traceback.print_exc()
        return None


# This allows you to test just this file by right-clicking and 'Run'
if __name__ == "__main__":
    result = process_and_get_html_and_file()
    if result:
        print("\n--- HTML PREVIEW ---")
        print(result[0][:500])