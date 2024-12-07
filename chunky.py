import pandas as pd
import os
from tqdm import tqdm  # Progress bar library

# Global Variables
FILE_PATH = 'data/emails_about_doors.xlsx'  # Path to the input Excel file
OUTPUT_DIR = 'emails_chunks_output'  # Directory to save the split files
ROWS_PER_CHUNK = 5000  # Number of rows per chunk
IGNORED_EMAILS = ['noreply@homesteadcabinet.net']  # List of known spam/advertisement email addresses
SPAM_KEYWORDS = ['promotion', 'sale', 'offer', 'unsubscribe', 'free', 'discount', 'advertisement', 'marketing']
SPAM_DOMAINS = ['.promo', '.info', 'marketing.com']  # Example domains often used for spam

# Global counter for filtered emails
emails_filtered_out = 0


def clean_body(body):
    """
    Cleans the email body by:
    1. Removing all quoted replies starting with "On [date] at [time] [sender] wrote:".
    2. Removing inline reply quotes (lines prefixed with '>').
    """
    if pd.isna(body) or not isinstance(body, str):
        return body

    # Remove inline quotes (lines starting with '>')
    body_lines = body.splitlines()
    cleaned_lines = [line for line in body_lines if not line.strip().startswith('>')]

    # Join the cleaned lines back into a single string
    cleaned_body = '\n'.join(cleaned_lines)

    return cleaned_body.strip()


def is_spam(row):
    """
    Determines if a row is spam based on known patterns and keywords.
    Handles non-string values gracefully.
    """
    from_email = str(row.get('From', '')).lower()  # Convert to string and lower case
    if any(from_email.endswith(domain) for domain in SPAM_DOMAINS):
        return True

    # Check subject and body for spam keywords
    subject = str(row.get('Subject', '')).lower()  # Convert to string and lower case
    body = str(row.get('Body', '')).lower()  # Convert to string and lower case
    if any(keyword in subject for keyword in SPAM_KEYWORDS) or any(keyword in body for keyword in SPAM_KEYWORDS):
        return True

    return False


def split_excel():
    """
    Splits an Excel file into smaller chunks based on the global variables.
    Filters out spam and advertisements, removes quoted replies, and includes a progress bar.
    """
    global emails_filtered_out

    # Load the Excel file
    excel_data = pd.ExcelFile(FILE_PATH)

    # Create the output directory if it doesn't exist
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # Process each sheet in the Excel file
    for sheet_name in excel_data.sheet_names:
        sheet_df = excel_data.parse(sheet_name)

        # Track original count
        original_count = len(sheet_df)

        # Remove rows where the 'From' column contains any email in IGNORED_EMAILS
        if 'From' in sheet_df.columns:
            sheet_df = sheet_df[~sheet_df['From'].isin(IGNORED_EMAILS)]

        # Filter out rows considered spam
        sheet_df = sheet_df[~sheet_df.apply(is_spam, axis=1)]

        # Clean email bodies
        if 'Body' in sheet_df.columns:
            sheet_df['Body'] = sheet_df['Body'].apply(clean_body)

        # Track filtered count
        filtered_count = original_count - len(sheet_df)
        emails_filtered_out += filtered_count

        num_rows = sheet_df.shape[0]
        chunk_count = 0

        # Initialize the progress bar
        with tqdm(total=num_rows, desc=f"Processing '{sheet_name}'", unit="rows") as pbar:
            # Split into chunks
            for start_row in range(0, num_rows, ROWS_PER_CHUNK):
                end_row = min(start_row + ROWS_PER_CHUNK, num_rows)
                chunk_df = sheet_df.iloc[start_row:end_row]
                output_file = os.path.join(OUTPUT_DIR, f"{sheet_name}_chunk_{chunk_count + 1}.xlsx")
                chunk_df.to_excel(output_file, index=False, sheet_name=sheet_name)
                chunk_count += 1
                pbar.update(end_row - start_row)

        print(f"Sheet '{sheet_name}' split into {chunk_count} chunks.")

    print(f"Files saved in: {OUTPUT_DIR}")
    print(f"Total emails filtered out: {emails_filtered_out}")


# Run the script
if __name__ == "__main__":
    split_excel()
