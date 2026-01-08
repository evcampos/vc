#!/usr/bin/env python3
"""
Email Organization Script
Moves non-priority messages to a "To be organized" folder/label
"""

import os
import sys
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pickle

# Gmail API scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

# Folder/Label name
ORGANIZE_FOLDER = "To be organized"

# Priority keywords (messages containing these will NOT be moved)
PRIORITY_KEYWORDS = [
    'urgent', 'important', 'asap', 'priority', 'critical',
    'meeting', 'invoice', 'payment', 'contract', 'signed'
]


def get_gmail_service():
    """Authenticate and return Gmail API service."""
    creds = None

    # Token file stores user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    # If no valid credentials, let user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists('credentials.json'):
                print("Error: credentials.json not found!")
                print("Please download OAuth 2.0 credentials from Google Cloud Console")
                sys.exit(1)

            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        # Save credentials for next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return build('gmail', 'v1', credentials=creds)


def get_or_create_label(service, label_name):
    """Get existing label or create new one if it doesn't exist."""
    try:
        # List all labels
        results = service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])

        # Check if label already exists
        for label in labels:
            if label['name'] == label_name:
                print(f"✓ Label '{label_name}' already exists")
                return label['id']

        # Create new label if it doesn't exist
        print(f"Creating label '{label_name}'...")
        label_object = {
            'name': label_name,
            'messageListVisibility': 'show',
            'labelListVisibility': 'labelShow'
        }

        created_label = service.users().labels().create(
            userId='me',
            body=label_object
        ).execute()

        print(f"✓ Label '{label_name}' created successfully")
        return created_label['id']

    except HttpError as error:
        print(f"An error occurred: {error}")
        sys.exit(1)


def is_priority_message(service, message_id):
    """Check if a message should be considered priority."""
    try:
        # Get message details
        message = service.users().messages().get(
            userId='me',
            id=message_id,
            format='metadata',
            metadataHeaders=['Subject', 'From']
        ).execute()

        # Check labels for IMPORTANT or STARRED
        labels = message.get('labelIds', [])
        if 'IMPORTANT' in labels or 'STARRED' in labels:
            return True

        # Check subject and from for priority keywords
        headers = message.get('payload', {}).get('headers', [])
        subject = ''
        from_email = ''

        for header in headers:
            if header['name'] == 'Subject':
                subject = header['value'].lower()
            elif header['name'] == 'From':
                from_email = header['value'].lower()

        # Check if any priority keyword is in subject or from
        for keyword in PRIORITY_KEYWORDS:
            if keyword in subject or keyword in from_email:
                return True

        return False

    except HttpError as error:
        print(f"Error checking message {message_id}: {error}")
        return True  # If error, assume it's priority to be safe


def move_messages_to_organize_folder(service, label_id, max_messages=100, dry_run=False):
    """Move non-priority messages to the organize folder."""
    try:
        # Get messages from inbox
        print(f"\nFetching messages from inbox...")
        results = service.users().messages().list(
            userId='me',
            labelIds=['INBOX'],
            maxResults=max_messages
        ).execute()

        messages = results.get('messages', [])

        if not messages:
            print("No messages found in inbox")
            return 0

        print(f"Found {len(messages)} messages in inbox")

        moved_count = 0
        skipped_count = 0

        for i, message in enumerate(messages, 1):
            message_id = message['id']

            # Check if message is priority
            if is_priority_message(service, message_id):
                skipped_count += 1
                continue

            # Move message to organize folder
            if not dry_run:
                try:
                    # Add the organize label and remove INBOX label
                    service.users().messages().modify(
                        userId='me',
                        id=message_id,
                        body={
                            'addLabelIds': [label_id],
                            'removeLabelIds': ['INBOX']
                        }
                    ).execute()
                    moved_count += 1
                    print(f"  [{i}/{len(messages)}] Moved message {message_id}")
                except HttpError as error:
                    print(f"  Error moving message {message_id}: {error}")
            else:
                moved_count += 1
                print(f"  [DRY RUN] Would move message {message_id}")

        return moved_count, skipped_count

    except HttpError as error:
        print(f"An error occurred: {error}")
        return 0, 0


def main():
    """Main function to organize emails."""
    print("=" * 60)
    print("Email Organization Tool")
    print("=" * 60)
    print(f"Folder: {ORGANIZE_FOLDER}")
    print(f"Priority keywords: {', '.join(PRIORITY_KEYWORDS)}")
    print("=" * 60)

    # Parse command line arguments
    dry_run = '--dry-run' in sys.argv
    max_messages = 100

    for arg in sys.argv:
        if arg.startswith('--max='):
            try:
                max_messages = int(arg.split('=')[1])
            except ValueError:
                print(f"Invalid max value: {arg}")
                sys.exit(1)

    if dry_run:
        print("\n*** DRY RUN MODE - No changes will be made ***\n")

    # Authenticate and get Gmail service
    print("Authenticating with Gmail...")
    service = get_gmail_service()
    print("✓ Authentication successful\n")

    # Get or create the organize label
    label_id = get_or_create_label(service, ORGANIZE_FOLDER)

    # Move messages
    moved_count, skipped_count = move_messages_to_organize_folder(
        service, label_id, max_messages, dry_run
    )

    # Print summary
    print("\n" + "=" * 60)
    print("Summary:")
    print(f"  Messages moved: {moved_count}")
    print(f"  Priority messages skipped: {skipped_count}")
    print("=" * 60)

    if dry_run:
        print("\nThis was a dry run. Use without --dry-run to actually move messages.")
    else:
        print(f"\n✓ Successfully organized emails!")


if __name__ == '__main__':
    main()
