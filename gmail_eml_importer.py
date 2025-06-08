# gmail_eml_importer.py
import os
import argparse
import base64
import email
import json
from email.message import Message
from email.utils import parsedate_tz, mktime_tz
from typing import Optional, List, Dict, Any
from tqdm import tqdm

# Google API imports
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Gmail API scopes needed for importing messages and managing labels
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

def authenticate_gmail(credentials_file: str) -> Any:
    """
    Authenticate with Gmail API using OAuth2.
    
    Args:
        credentials_file: Path to the credentials.json file from Google Cloud Console
        
    Returns:
        Authenticated Gmail service object
    """
    creds = None
    token_file = 'token.json'
    
    # Load existing token if available
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)
    
    # If there are no valid credentials, authenticate
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(credentials_file):
                raise FileNotFoundError(f"Credentials file not found: {credentials_file}")
            
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save credentials for next run
        with open(token_file, 'w') as token:
            token.write(creds.to_json())
    
    return build('gmail', 'v1', credentials=creds)

def get_or_create_label(service: Any, label_name: str) -> Optional[str]:
    """
    Get the ID of a Gmail label, creating it if it doesn't exist.
    
    Args:
        service: Authenticated Gmail service object
        label_name: Name of the label to find or create
        
    Returns:
        Label ID if successful, None otherwise
    """
    try:
        # List existing labels
        results = service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])
        
        # Check if label already exists
        for label in labels:
            if label['name'] == label_name:
                return label['id']
        
        # Create new label if it doesn't exist
        label_object = {
            'name': label_name,
            'labelListVisibility': 'labelShow',
            'messageListVisibility': 'show'
        }
        
        created_label = service.users().labels().create(
            userId='me', 
            body=label_object
        ).execute()
        
        print(f"  -> Created new label: '{label_name}'")
        return created_label['id']
        
    except HttpError as error:
        print(f"  [ERROR] Failed to get/create label '{label_name}': {error}")
        return None

def message_exists(service: Any, message_id_header: str) -> bool:
    """
    Check if a message with the given Message-ID already exists in Gmail.
    
    Args:
        service: Authenticated Gmail service object
        message_id_header: The Message-ID header value to search for
        
    Returns:
        True if message exists, False otherwise
    """
    try:
        # Clean the Message-ID for search (remove angle brackets if present)
        clean_msgid = message_id_header.strip('<>')
        
        # Search for messages with this Message-ID
        query = f'rfc822msgid:{clean_msgid}'
        results = service.users().messages().list(userId='me', q=query, maxResults=1).execute()
        
        messages = results.get('messages', [])
        return len(messages) > 0
        
    except HttpError as error:
        # If search fails, assume message doesn't exist to avoid skipping imports
        print(f"  [WARNING] Could not check for duplicate (will import anyway): {error}")
        return False
    except Exception:
        # For any other error, assume message doesn't exist
        return False

def import_eml_to_gmail(service: Any, eml_path: str, label_name: Optional[str] = None, check_duplicates: bool = True) -> tuple[bool, str]:
    """
    Import a single .eml file to Gmail using the Gmail API.
    
    Args:
        service: Authenticated Gmail service object
        eml_path: Path to the .eml file
        label_name: Optional label to apply to the imported message
        check_duplicates: Whether to check for existing messages with same Message-ID
        
    Returns:
        Tuple of (success: bool, status_message: str)
    """
    try:
        basename = os.path.basename(eml_path)
        
        # Read the .eml file
        with open(eml_path, 'rb') as f:
            eml_bytes = f.read()
        
        # Parse the message to extract information
        msg: Message = email.message_from_bytes(eml_bytes)
        
        # Check for duplicates if requested
        if check_duplicates:
            message_id_header = msg.get('Message-ID')
            if message_id_header:
                if message_exists(service, message_id_header):
                    return True, f"SKIPPED (duplicate): {basename}"
        
        # Prepare the message for Gmail API
        message_body = {
            'raw': base64.urlsafe_b64encode(eml_bytes).decode('utf-8')
        }
        
        # Try to preserve the original date
        date_header = msg.get('Date')
        if date_header:
            try:
                date_tuple = parsedate_tz(date_header)
                if date_tuple:
                    # Convert to milliseconds since epoch (Gmail API format)
                    utc_timestamp = mktime_tz(date_tuple)
                    internal_date = str(int(utc_timestamp * 1000))
                    message_body['internalDate'] = internal_date
            except (ValueError, TypeError):
                # If date parsing fails, let Gmail use current time
                pass
        
        # Import the message
        imported_message = service.users().messages().import_(
            userId='me',
            body=message_body,
            neverMarkSpam=True,
            processForCalendar=False
        ).execute()
        
        message_id = imported_message['id']
        
        # Apply label if specified
        if label_name:
            label_id = get_or_create_label(service, label_name)
            if label_id:
                try:
                    service.users().messages().modify(
                        userId='me',
                        id=message_id,
                        body={'addLabelIds': [label_id]}
                    ).execute()
                    return True, f"IMPORTED with label: {basename}"
                except HttpError as error:
                    return True, f"IMPORTED (label failed): {basename}"
            else:
                return True, f"IMPORTED (no label): {basename}"
        else:
            return True, f"IMPORTED: {basename}"
        
    except HttpError as error:
        return False, f"FAILED: {basename} - {error}"
    except Exception as e:
        return False, f"ERROR: {basename} - {e}"

def main():
    """Main function to parse arguments and orchestrate the import process."""
    parser = argparse.ArgumentParser(
        description="Import .eml files into Gmail using the Gmail API, preserving dates and applying labels.",
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("path", help="Path to a single .eml file or a directory containing .eml files.")
    parser.add_argument("-c", "--credentials", default="credentials.json", 
                       help="Path to the Gmail API credentials file (default: credentials.json).")
    parser.add_argument("-l", "--label", help="Gmail label to apply to imported messages.")
    parser.add_argument("-r", "--recursive", action="store_true", 
                       help="Search for .eml files recursively in the given directory.")
    parser.add_argument("--no-duplicates", action="store_true",
                       help="Skip duplicate checking (faster but may create duplicates).")
    
    args = parser.parse_args()

    # Check if credentials file exists
    if not os.path.exists(args.credentials):
        print(f"Error: Credentials file not found: {args.credentials}")
        print("\nTo get Gmail API credentials:")
        print("1. Go to https://console.cloud.google.com/")
        print("2. Create a new project or select existing one")
        print("3. Enable Gmail API")
        print("4. Create OAuth2 credentials (Desktop application)")
        print("5. Download the credentials.json file")
        return

    # Find .eml files
    eml_files = []
    if os.path.isfile(args.path) and args.path.lower().endswith('.eml'):
        eml_files.append(args.path)
    elif os.path.isdir(args.path):
        if args.recursive:
            for root, _, files in os.walk(args.path):
                for file in files:
                    if file.lower().endswith('.eml'):
                        eml_files.append(os.path.join(root, file))
        else:
            for file in os.listdir(args.path):
                if file.lower().endswith('.eml'):
                    eml_files.append(os.path.join(args.path, file))
    
    if not eml_files:
        print(f"No .eml files found in the specified path: {args.path}")
        return

    print(f"Found {len(eml_files)} .eml file(s) to import.")

    try:
        print("Authenticating with Gmail API...")
        service = authenticate_gmail(args.credentials)
        
        print(f"\nStarting import process...")
        if not args.no_duplicates:
            print("(Checking for duplicates - use --no-duplicates to skip)")
        print()
        
        # Counters for statistics
        successful_imports = 0
        failed_imports = 0
        skipped_duplicates = 0
        
        # Process files with progress bar
        with tqdm(total=len(eml_files), desc="Importing", unit="file") as pbar:
            for eml_file in sorted(eml_files):
                success, status_msg = import_eml_to_gmail(
                    service, 
                    eml_file, 
                    args.label, 
                    check_duplicates=not args.no_duplicates
                )
                
                # Update progress bar with current file info
                pbar.set_postfix_str(status_msg)
                pbar.update(1)
                
                # Update counters
                if success:
                    if "SKIPPED" in status_msg:
                        skipped_duplicates += 1
                    else:
                        successful_imports += 1
                else:
                    failed_imports += 1
        
        print(f"\nImport process finished.")
        print(f"Successfully imported: {successful_imports}")
        if not args.no_duplicates:
            print(f"Skipped duplicates: {skipped_duplicates}")
        print(f"Failed imports: {failed_imports}")
        
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    main()