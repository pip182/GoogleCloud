import os
import mailbox
import csv
from email.utils import parsedate_to_datetime
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from datetime import datetime
from tqdm import tqdm  # For progress bar


# Configuration
current_directory = r"D:\DataDump"
keywords = ['door', 'door-style', 'door shop', 'door procedure', 'problem', 'issue',
            'wrong', 'incorrect', 'mistake', 'error', 'fault', 'defect', 'flaw', 'bug',
            'trouble', 'difficulty', 'complication', 'dilemma', 'predicament',
            'quandary', 'plight', 'obstacle', 'hurdle', 'barrier', 'impediment', 'snag',
            'hiccup', 'setback', 'disadvantage', 'weakness', 'shortcoming', 'deficiency',
            'failing', 'imperfection', 'blemish']
chunk_size_gb = 2  # Size of each chunk in GB
failed_items_csv = os.path.join(current_directory, "failed_messages.csv")
door_emails_file = os.path.join(current_directory, "emails_about_doors_and_issues.xlsx")

# Spam filtering configuration
spam_keywords = [
    "unsubscribe", "promo", "free", "offer", "sale", "discount",
    "marketing", "ad", "advertisement", "newsletter", "click here", "limited-time deal", "exclusive offer"
]  # Extend this list as needed

ignored_senders = [
    "noreply@homesteadcabinet.net",  # Add other ignored senders here
    "mailer-daemon@googlemail.com",
]


def parse_email_date(email_date):
    """Parse the email date to extract year, month, and reformat it for Excel."""
    try:
        # Handle common date parsing using email.utils
        parsed_date = parsedate_to_datetime(email_date)
        # Return the date in Excel-compatible format (YYYY-MM-DD HH:MM:SS)
        return parsed_date.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return "Unknown"


def is_spam_or_advertisement(email_from, subject, body):
    """
    Determines if an email is likely spam or an advertisement.

    Args:
        email_from (str): Sender email address.
        subject (str): Email subject line.
        body (str): Email body content.

    Returns:
        bool: True if the email is spam or an advertisement; False otherwise.
    """
    # Check if the sender is blocked
    if email_from.lower() in blocked_senders:
        return True

    # Check for spam keywords in subject or body
    subject_lower = subject.lower() if subject else ""
    body_lower = body.lower() if body else ""
    for keyword in spam_keywords:
        if keyword in subject_lower or keyword in body_lower:
            return True

    return False


def remove_quoted_lines(body):
    """
    Remove quoted lines from the email body (lines that start with '>').

    Args:
        body (str): The email body.

    Returns:
        str: The email body with quoted lines removed.
    """
    if not body:
        return body

    lines = body.splitlines()
    filtered_lines = [line for line in lines if not line.strip().startswith('>')]
    return "\n".join(filtered_lines)


def get_or_create_workbook(file_path, workbooks):
    """Get or create an openpyxl Workbook for .xlsx files."""
    if file_path not in workbooks:
        wb = Workbook()
        sheet = wb.active
        sheet.title = "Emails"
        # Append header row with bold formatting
        header = ["From", "Subject", "Date", "Body"]
        sheet.append(header)
        for col in sheet[1]:
            col.font = Font(bold=True)  # Make the header bold
            col.alignment = Alignment(horizontal="left", vertical="top")  # Align to top-left
        workbooks[file_path] = {"workbook": wb, "sheet": sheet}
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


def split_mbox_by_size(input_mbox, output_dir, max_size_gb):
    """
    Splits a large mbox file into smaller chunks based on file size.

    Args:
        input_mbox (str): Path to the input mbox file.
        output_dir (str): Directory to save the output chunked mbox files.
        max_size_gb (int): Maximum size of each chunk in GB.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    print("BLOWING CHUNKS.... PLEASE WAIT....")

    max_chunk_size = max_size_gb * 1024 * 1024 * 1024  # Convert GB to bytes
    chunk_index = 0
    current_chunk_size = 0
    chunk_file = None

    with open(input_mbox, 'rb') as infile:
        while True:
            chunk_path = os.path.join(output_dir, f"chunk_{chunk_index}.mbox")
            chunk_file = open(chunk_path, 'wb')
            print(f"Creating chunk: {chunk_path}")

            # Progress bar for each chunk
            with tqdm(total=max_chunk_size, desc=f"Chunk {chunk_index}", unit="B", unit_scale=True, unit_divisor=1024) as pbar:
                current_chunk_size = 0
                while current_chunk_size < max_chunk_size:
                    line = infile.readline()
                    if not line:  # End of file
                        break
                    chunk_file.write(line)
                    current_chunk_size += len(line)
                    pbar.update(len(line))

            chunk_file.close()
            chunk_index += 1

            # Stop if we've reached the end of the file
            if not line:  # No more data to read
                break

    print(f"Splitting complete. Created {chunk_index} chunks in {output_dir}")
    return [os.path.join(output_dir, f"chunk_{i}.mbox") for i in range(chunk_index)]


def validate_chunks(chunk_dir, expected_chunk_count):
    """
    Validates that all expected chunks exist in the chunk directory.

    Args:
        chunk_dir (str): Directory containing chunked mbox files.
        expected_chunk_count (int): Number of expected chunks.

    Returns:
        bool: True if all expected chunks exist, False otherwise.
    """
    chunk_files = [f for f in os.listdir(chunk_dir) if f.startswith("chunk_") and f.endswith(".mbox")]
    if len(chunk_files) == expected_chunk_count:
        return True
    print(f"Missing chunks detected in {chunk_dir}. Re-creating chunks.")
    return False


def save_to_csv(csv_path, data_row):
    """Safely save a row of data to the CSV file."""
    try:
        with open(csv_path, mode='a', newline='', encoding='utf-8') as f:
            csv_writer = csv.writer(f)
            csv_writer.writerow(data_row)
    except Exception as csv_error:
        print(f"ERROR: Failed to save row to CSV {csv_path}. Data: {data_row}. Error: {csv_error}")


def clear_workbooks(workbooks, preserve_file=None):
    """
    Clears all workbooks from memory except the one to preserve.

    Args:
        workbooks (dict): Dictionary of workbooks.
        preserve_file (str): File path of the workbook to preserve.
    """
    files_to_clear = list(workbooks.keys())
    for file_path in files_to_clear:
        if file_path != preserve_file:
            del workbooks[file_path]
    print("Cleared all workbooks except:", preserve_file)


def mbox_to_excel_stream_grouped(mbox_file, root_name, keywords, workbooks, failed_csv_path):
    """Convert a single mbox file to multiple Excel files grouped by year and month."""
    try:
        # Open the mbox file
        mbox = mailbox.mbox(mbox_file)
        total_messages = max(mbox.keys()) + 1  # Faster way to determine total messages

        # Determine output directory (same as mbox file)
        output_dir = os.path.dirname(mbox_file)

        # Initialize progress bar
        with tqdm(total=total_messages, desc=f"Processing {os.path.basename(mbox_file)}", unit="email") as pbar:
            # Process emails
            for i, message in enumerate(mbox):
                try:
                    # Extract essential headers
                    email_from = message.get('From', '').strip().lower()
                    subject = message.get('Subject', '').strip()
                    date = message.get('Date', '').strip()

                    # Convert date to Excel-compatible format
                    formatted_date = parse_email_date(date)

                    # Parse year and month for grouping
                    year, month = formatted_date.split('-')[0], formatted_date.split('-')[1] if formatted_date != "Unknown" else ("Unknown", "Unknown")

                    # Handle "Unknown" year/month
                    if year == "Unknown" or month == "Unknown":
                        key = "Unknown"
                    else:
                        key = f"{year}-{month}"

                    # Extract plain text body only
                    body = None
                    if message.is_multipart():
                        for part in message.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition", ""))
                            # Only process plain text parts, skip attachments
                            if content_type == "text/plain" and "attachment" not in content_disposition:
                                body = part.get_payload(decode=True).decode(errors='ignore')
                                break
                    else:
                        body = message.get_payload(decode=True).decode(errors='ignore')

                    # If no body is found, skip this email
                    if not body:
                        pbar.update(1)
                        continue

                    # Remove quoted lines from the email body
                    body = remove_quoted_lines(body)

                    # Filter out spam/advertisement emails
                    if is_spam_or_advertisement(email_from, subject, body):
                        pbar.update(1)
                        continue

                    # Check for keywords in the email's subject or body
                    if contains_keywords(subject, keywords) or contains_keywords(body, keywords):
                        door_data = get_or_create_workbook(door_emails_file, workbooks)
                        sheet = door_data["sheet"]
                        sheet.append([email_from, subject, formatted_date, body])

                    # Define the output file path based on year and month
                    output_file = os.path.join(output_dir, f"{root_name}_{key}.xlsx")

                    # Get or create workbook
                    data = get_or_create_workbook(output_file, workbooks)
                    sheet = data["sheet"]
                    sheet.append([email_from, subject, formatted_date, body])

                    # Update progress bar
                    pbar.update(1)

                except Exception as e:
                    # Handle errors by saving to the CSV file
                    save_to_csv(failed_csv_path, [email_from, subject, formatted_date, body])
                    print(f"Failed to process message {i} in {mbox_file}. Error saved to CSV. Error: {e}")

        # Final save for all workbooks
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

    workbooks = {}  # Dictionary to track open workbooks

    for root_dir in root_dirs:
        root_name = os.path.basename(root_dir)  # Use root directory name for file naming
        print(f"Processing root directory: {root_name}")

        for root, _, files in os.walk(root_dir):
            for file in files:
                if file.endswith('.mbox') and "chunks" not in root:  # Ignore chunks directory
                    mbox_file = os.path.join(root, file)

                    # Chunk directory
                    chunk_dir = os.path.join(root, "chunks")

                    # Determine expected chunk count
                    total_size = os.path.getsize(mbox_file)
                    expected_chunk_count = (total_size // (chunk_size_gb * 1024 * 1024 * 1024)) + 1

                    # Check if chunks exist and are valid
                    if os.path.exists(chunk_dir) and validate_chunks(chunk_dir, expected_chunk_count):
                        print(f"Valid chunks already exist for {mbox_file}. Skipping chunk creation.")
                        chunks = [os.path.join(chunk_dir, f"chunk_{i}.mbox") for i in range(expected_chunk_count)]
                    else:
                        print(f"Re-chunking {mbox_file} into {chunk_dir}.")
                        chunks = split_mbox_by_size(mbox_file, chunk_dir, chunk_size_gb)

                    # Process each chunk
                    for chunk in chunks:
                        print(f"Processing chunk: {chunk}")
                        mbox_to_excel_stream_grouped(chunk, root_name, keywords, workbooks, failed_items_csv)

        # Clear workbooks except for the emails_about_doors workbook
        clear_workbooks(workbooks, preserve_file=door_emails_file)

    print("Processing complete.")


if __name__ == '__main__':
    if os.path.exists(current_directory):
        # Initialize the failed items CSV with a header
        with open(failed_items_csv, mode='w', newline='', encoding='utf-8') as f:
            csv_writer = csv.writer(f)
            csv_writer.writerow(["From", "Subject", "Date", "Body"])

        find_and_convert_mbox_files(current_directory)
    else:
        print(f"Error: Directory {current_directory} does not exist.")
