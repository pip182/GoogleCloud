import os
import mailbox
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from datetime import datetime
from tqdm import tqdm  # For progress bar


# Configuration
current_directory = r"D:\DataDump"
keywords = ['door', 'door-style', 'door shop', 'door procedure']


def parse_email_date(email_date):
    """Parse the email date to extract year and month."""
    try:
        parsed_date = datetime.strptime(email_date, '%a, %d %b %Y %H:%M:%S %z')
        return parsed_date.year, parsed_date.month
    except Exception:
        return "Unknown", "Unknown"


def get_or_create_workbook(file_path, workbooks):
    """Get or create an openpyxl Workbook for .xlsx files."""
    if file_path not in workbooks:
        wb = Workbook()
        sheet = wb.active
        sheet.title = "Emails"
        # Append header row with bold formatting
        header = ["From", "To", "Subject", "Date", "Body"]
        sheet.append(header)
        for col in sheet[1]:
            col.font = Font(bold=True)  # Make the header bold
            col.alignment = Alignment(horizontal="left", vertical="top")  # Align to top-left
        workbooks[file_path] = {"workbook": wb, "sheet": sheet, "row": 2}  # Start at row 2
    return workbooks[file_path]


def contains_keywords(text, keywords):
    """Check if the text contains any of the keywords."""
    text_lower = text.lower()
    return any(keyword in text_lower for keyword in keywords)


def save_workbooks(workbooks):
    """Save all open workbooks and adjust row heights."""
    for file_path, data in workbooks.items():
        sheet = data["sheet"]
        for row in sheet.iter_rows():
            # Set each row height to 0.22
            sheet.row_dimensions[row[0].row].height = 0.22 * 72  # Approximate height in points
            for cell in row:
                # Ensure all text is aligned top-left
                cell.alignment = Alignment(horizontal="left", vertical="top")
        data["workbook"].save(file_path)


def mbox_to_excel_stream_grouped(mbox_file, root_name, keywords, door_emails_file, workbooks):
    """Convert a single mbox file to multiple Excel files grouped by year and month."""
    try:
        # Open the mbox file
        mbox = mailbox.mbox(mbox_file)
        total_messages = max(mbox.keys()) + 1  # Faster way to determine total messages

        # Determine output directory (same as mbox file)
        output_dir = os.path.dirname(mbox_file)

        message_count = 0

        # Initialize progress bar
        with tqdm(total=total_messages, desc="Processing Emails", unit="email") as pbar:
            # Process emails
            for i, message in enumerate(mbox):
                try:
                    # Extract email fields
                    email_from = message.get('From', '')
                    email_to = message.get('To', '')
                    subject = message.get('Subject', '')
                    date = message.get('Date', '')

                    # Parse year and month
                    year, month = parse_email_date(date)

                    # Handle "Unknown" year/month
                    if year == "Unknown" or month == "Unknown":
                        key = "Unknown"
                    else:
                        key = f"{year}-{int(month):02d}"

                    # Extract email body
                    if message.is_multipart():
                        body = ''.join(
                            part.get_payload(decode=True).decode(errors='ignore')
                            for part in message.walk()
                            if part.get_content_type() == 'text/plain'
                        )
                    else:
                        body = message.get_payload(decode=True).decode(errors='ignore')

                    # Check for keywords in the email's subject or body
                    if contains_keywords(subject, keywords) or contains_keywords(body, keywords):
                        door_data = get_or_create_workbook(door_emails_file, workbooks)
                        sheet = door_data["sheet"]
                        sheet.append([email_from, email_to, subject, date, body])

                    # Define the output file path based on year and month
                    output_file = os.path.join(output_dir, f"{root_name}_{key}.xlsx")

                    # Get or create workbook
                    data = get_or_create_workbook(output_file, workbooks)
                    sheet = data["sheet"]
                    sheet.append([email_from, email_to, subject, date, body])
                    message_count += 1

                    # Update progress bar
                    pbar.update(1)

                    # Save workbooks every 50 messages
                    if message_count % 50 == 0 or message_count == total_messages:
                        save_workbooks(workbooks)

                except Exception as e:
                    print(f"Failed to process message {i} in {mbox_file}: {e}")

        # Final save for any unsaved data
        save_workbooks(workbooks)

        print(f"SUCCESS: Processed and grouped emails from {mbox_file}")

    except Exception as e:
        print(f"ERROR: Failed to process {mbox_file}: {e}")


def find_and_convert_mbox_files(start_dir):
    """Recursively find and convert all mbox files in each root directory."""
    print(f"Scanning directory: {start_dir}")
    root_dirs = [os.path.join(start_dir, d) for d in os.listdir(start_dir) if os.path.isdir(os.path.join(start_dir, d))]

    if not root_dirs:
        print("No root directories found.")
        return

    door_emails_file = os.path.join(start_dir, "emails_about_doors.xlsx")
    workbooks = {}  # Dictionary to track open workbooks

    for root_dir in root_dirs:
        root_name = os.path.basename(root_dir)  # Use root directory name for file naming
        print(f"Processing root directory: {root_name}")

        for root, _, files in os.walk(root_dir):
            for file in files:
                if file.endswith('.mbox'):
                    mbox_file = os.path.join(root, file)

                    # Log the file found
                    print(f"Found: {mbox_file}")

                    # Convert the mbox file to grouped Excel files
                    mbox_to_excel_stream_grouped(mbox_file, root_name, keywords, door_emails_file, workbooks)

    print("Processing complete.")


if __name__ == '__main__':
    if os.path.exists(current_directory):
        find_and_convert_mbox_files(current_directory)
    else:
        print(f"Error: Directory {current_directory} does not exist.")
