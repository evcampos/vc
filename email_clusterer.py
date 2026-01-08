#!/usr/bin/env python3
"""
Email Clustering System for macOS Apple Mail
Automatically clusters emails by themes and checks calendar relevance
"""

import subprocess
import json
import re
from datetime import datetime, timedelta
from pathlib import Path
import sys
from typing import List, Dict, Tuple, Optional
import os

try:
    import pandas as pd
    import openpyxl
except ImportError:
    print("ERROR: Required packages not installed.")
    print("Please run: pip3 install pandas openpyxl")
    sys.exit(1)


class EmailClusterer:
    """Main class for email clustering and calendar integration"""

    def __init__(self, database_path: str = None):
        """Initialize the email clusterer with database path"""
        if database_path is None:
            # Default to user's Documents folder
            home = Path.home()
            database_path = home / "Documents" / "EmailClusterDatabase.xlsx"

        self.database_path = Path(database_path)
        self.categories = {}
        self.keyword_mappings = {}
        self.load_or_create_database()

    def load_or_create_database(self):
        """Load existing database or create new one"""
        if self.database_path.exists():
            print(f"Loading database from: {self.database_path}")
            self.load_database()
        else:
            print(f"Creating new database at: {self.database_path}")
            self.create_database()

    def create_database(self):
        """Create a new Excel database with default structure"""
        # Create default categories
        default_categories = {
            'Work': ['meeting', 'project', 'deadline', 'report', 'presentation'],
            'Personal': ['family', 'friend', 'dinner', 'appointment'],
            'Finance': ['invoice', 'payment', 'bill', 'receipt', 'transaction'],
            'Shopping': ['order', 'delivery', 'shipment', 'purchase', 'cart'],
            'Social': ['event', 'invitation', 'party', 'gathering'],
            'Travel': ['flight', 'hotel', 'booking', 'reservation', 'trip'],
            'Newsletter': ['newsletter', 'digest', 'update', 'subscribe'],
            'Support': ['ticket', 'support', 'help', 'issue', 'problem'],
            'Uncategorized': []
        }

        # Create DataFrames
        categories_data = []
        for category, keywords in default_categories.items():
            for keyword in keywords:
                categories_data.append({
                    'Category': category,
                    'Keyword': keyword,
                    'Active': True,
                    'Created': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                })

        df_categories = pd.DataFrame(categories_data)
        df_logs = pd.DataFrame(columns=[
            'Timestamp', 'Subject', 'Sender', 'Category',
            'CalendarMatch', 'MatchedEvent', 'Confidence'
        ])
        df_stats = pd.DataFrame(columns=[
            'Date', 'TotalEmails', 'Categorized', 'WithCalendarMatch'
        ])

        # Save to Excel
        with pd.ExcelWriter(self.database_path, engine='openpyxl') as writer:
            df_categories.to_excel(writer, sheet_name='Categories', index=False)
            df_logs.to_excel(writer, sheet_name='EmailLogs', index=False)
            df_stats.to_excel(writer, sheet_name='Statistics', index=False)

        print(f"✓ Database created with {len(default_categories)} default categories")
        self.load_database()

    def load_database(self):
        """Load categories and keyword mappings from database"""
        try:
            df = pd.read_excel(self.database_path, sheet_name='Categories')

            # Build keyword mappings
            for _, row in df.iterrows():
                if row['Active']:
                    category = row['Category']
                    keyword = str(row['Keyword']).lower()

                    if category not in self.categories:
                        self.categories[category] = []

                    self.categories[category].append(keyword)
                    self.keyword_mappings[keyword] = category

            print(f"✓ Loaded {len(self.categories)} categories with {len(self.keyword_mappings)} keywords")
        except Exception as e:
            print(f"Error loading database: {e}")
            self.create_database()

    def save_log_entry(self, email_data: Dict):
        """Save email processing log to database"""
        try:
            # Read existing logs
            df_logs = pd.read_excel(self.database_path, sheet_name='EmailLogs')

            # Append new entry
            new_entry = pd.DataFrame([{
                'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                'Subject': email_data.get('subject', ''),
                'Sender': email_data.get('sender', ''),
                'Category': email_data.get('category', 'Uncategorized'),
                'CalendarMatch': email_data.get('calendar_match', False),
                'MatchedEvent': email_data.get('matched_event', ''),
                'Confidence': email_data.get('confidence', 0)
            }])

            df_logs = pd.concat([df_logs, new_entry], ignore_index=True)

            # Save back to Excel
            with pd.ExcelWriter(self.database_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                df_logs.to_excel(writer, sheet_name='EmailLogs', index=False)

        except Exception as e:
            print(f"Warning: Could not save log entry: {e}")

    def update_statistics(self, total_emails: int, categorized: int, calendar_matches: int):
        """Update daily statistics"""
        try:
            df_stats = pd.read_excel(self.database_path, sheet_name='Statistics')

            today = datetime.now().strftime('%Y-%m-%d')

            # Update or create today's entry
            if today in df_stats['Date'].values:
                idx = df_stats[df_stats['Date'] == today].index[0]
                df_stats.at[idx, 'TotalEmails'] = total_emails
                df_stats.at[idx, 'Categorized'] = categorized
                df_stats.at[idx, 'WithCalendarMatch'] = calendar_matches
            else:
                new_stat = pd.DataFrame([{
                    'Date': today,
                    'TotalEmails': total_emails,
                    'Categorized': categorized,
                    'WithCalendarMatch': calendar_matches
                }])
                df_stats = pd.concat([df_stats, new_stat], ignore_index=True)

            with pd.ExcelWriter(self.database_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                df_stats.to_excel(writer, sheet_name='Statistics', index=False)

        except Exception as e:
            print(f"Warning: Could not update statistics: {e}")

    def get_inbox_emails(self, limit: int = 50) -> List[Dict]:
        """Fetch recent emails from Apple Mail inbox using AppleScript"""
        applescript = f'''
        tell application "Mail"
            set emailList to {{}}
            set inboxMessages to messages of inbox whose read status is false

            repeat with i from 1 to (count of inboxMessages)
                if i > {limit} then exit repeat

                set theMessage to item i of inboxMessages
                set emailData to {{}}
                set emailData to emailData & {{subject:(subject of theMessage)}}
                set emailData to emailData & {{sender:(sender of theMessage)}}
                set emailData to emailData & {{dateReceived:(date received of theMessage as string)}}
                set emailData to emailData & {{messageId:(id of theMessage as string)}}

                set end of emailList to emailData
            end repeat

            return emailList
        end tell
        '''

        try:
            result = subprocess.run(
                ['osascript', '-e', applescript],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode != 0:
                print(f"AppleScript error: {result.stderr}")
                return []

            # Parse AppleScript output
            emails = self.parse_applescript_list(result.stdout)
            return emails

        except Exception as e:
            print(f"Error fetching emails: {e}")
            return []

    def parse_applescript_list(self, output: str) -> List[Dict]:
        """Parse AppleScript list output into Python dictionaries"""
        emails = []

        # Simple parsing of AppleScript record format
        # Format: {subject:"...", sender:"...", dateReceived:"...", messageId:"..."}

        records = re.findall(r'\{subject:(.*?), sender:(.*?), dateReceived:(.*?), messageId:(.*?)\}', output, re.DOTALL)

        for record in records:
            email = {
                'subject': record[0].strip('"'),
                'sender': record[1].strip('"'),
                'date_received': record[2].strip('"'),
                'message_id': record[3].strip('"')
            }
            emails.append(email)

        return emails

    def get_calendar_events(self, days_ahead: int = 7) -> List[Dict]:
        """Fetch upcoming calendar events using AppleScript"""
        start_date = datetime.now()
        end_date = start_date + timedelta(days=days_ahead)

        applescript = f'''
        tell application "Calendar"
            set startDate to (current date)
            set endDate to startDate + ({days_ahead} * days)

            set eventList to {{}}
            set allCalendars to every calendar

            repeat with cal in allCalendars
                set calEvents to (every event of cal whose start date ≥ startDate and start date ≤ endDate)

                repeat with evt in calEvents
                    set eventData to {{}}
                    set eventData to eventData & {{summary:(summary of evt)}}
                    set eventData to eventData & {{startDate:(start date of evt as string)}}
                    set eventData to eventData & {{location:(location of evt)}}

                    set end of eventList to eventData
                end repeat
            end repeat

            return eventList
        end tell
        '''

        try:
            result = subprocess.run(
                ['osascript', '-e', applescript],
                capture_output=True,
                text=True,
                timeout=30
            )

            if result.returncode != 0:
                print(f"Calendar access error: {result.stderr}")
                return []

            events = self.parse_calendar_events(result.stdout)
            return events

        except Exception as e:
            print(f"Error fetching calendar events: {e}")
            return []

    def parse_calendar_events(self, output: str) -> List[Dict]:
        """Parse calendar events from AppleScript output"""
        events = []

        records = re.findall(r'\{summary:(.*?), startDate:(.*?), location:(.*?)\}', output, re.DOTALL)

        for record in records:
            event = {
                'summary': record[0].strip('"'),
                'start_date': record[1].strip('"'),
                'location': record[2].strip('"')
            }
            events.append(event)

        return events

    def categorize_email(self, email: Dict) -> Tuple[str, float]:
        """Categorize an email based on subject and learned patterns"""
        subject = email.get('subject', '').lower()
        sender = email.get('sender', '').lower()

        # Combined text for analysis
        text = f"{subject} {sender}"

        # Score each category
        category_scores = {}

        for category, keywords in self.categories.items():
            score = 0
            matched_keywords = []

            for keyword in keywords:
                if keyword in text:
                    score += 1
                    matched_keywords.append(keyword)

            if score > 0:
                category_scores[category] = {
                    'score': score,
                    'keywords': matched_keywords
                }

        # Determine best category
        if category_scores:
            best_category = max(category_scores.items(), key=lambda x: x[1]['score'])
            confidence = min(best_category[1]['score'] / 3.0, 1.0)  # Normalize confidence
            return best_category[0], confidence

        return 'Uncategorized', 0.0

    def check_calendar_match(self, email: Dict, events: List[Dict]) -> Tuple[bool, str]:
        """Check if email subject relates to any calendar event"""
        subject = email.get('subject', '').lower()
        subject_words = set(re.findall(r'\w+', subject))

        best_match = None
        best_score = 0

        for event in events:
            event_summary = event.get('summary', '').lower()
            event_words = set(re.findall(r'\w+', event_summary))

            # Calculate word overlap
            common_words = subject_words & event_words

            # Filter out common words
            common_words = {w for w in common_words if len(w) > 3}

            if len(common_words) > best_score:
                best_score = len(common_words)
                best_match = event.get('summary', '')

        # Threshold for considering a match
        if best_score >= 2:
            return True, best_match

        return False, ''

    def process_emails(self, limit: int = 50):
        """Main processing function"""
        print("\n" + "="*60)
        print("EMAIL CLUSTERING SYSTEM")
        print("="*60)

        # Fetch emails
        print("\n[1/4] Fetching emails from Apple Mail...")
        emails = self.get_inbox_emails(limit)
        print(f"✓ Found {len(emails)} unread emails")

        if not emails:
            print("No unread emails to process.")
            return

        # Fetch calendar events
        print("\n[2/4] Fetching calendar events...")
        events = self.get_calendar_events(days_ahead=14)
        print(f"✓ Found {len(events)} upcoming events")

        # Process each email
        print("\n[3/4] Processing and categorizing emails...")
        print("-" * 60)

        categorized_count = 0
        calendar_match_count = 0

        for idx, email in enumerate(emails, 1):
            subject = email.get('subject', 'No Subject')
            sender = email.get('sender', 'Unknown')

            # Categorize
            category, confidence = self.categorize_email(email)
            if category != 'Uncategorized':
                categorized_count += 1

            # Check calendar
            has_calendar_match, matched_event = self.check_calendar_match(email, events)
            if has_calendar_match:
                calendar_match_count += 1

            # Display result
            print(f"\n[{idx}/{len(emails)}] {subject[:50]}")
            print(f"    From: {sender[:40]}")
            print(f"    Category: {category} (confidence: {confidence:.2f})")

            if has_calendar_match:
                print(f"    📅 Calendar Match: {matched_event}")

            # Save log
            email_data = {
                'subject': subject,
                'sender': sender,
                'category': category,
                'confidence': confidence,
                'calendar_match': has_calendar_match,
                'matched_event': matched_event
            }
            self.save_log_entry(email_data)

        # Update statistics
        print("\n[4/4] Updating statistics...")
        self.update_statistics(len(emails), categorized_count, calendar_match_count)

        # Summary
        print("\n" + "="*60)
        print("SUMMARY")
        print("="*60)
        print(f"Total Emails Processed: {len(emails)}")
        print(f"Successfully Categorized: {categorized_count} ({categorized_count/len(emails)*100:.1f}%)")
        print(f"Calendar Matches Found: {calendar_match_count} ({calendar_match_count/len(emails)*100:.1f}%)")
        print(f"\nDatabase: {self.database_path}")
        print("="*60)


def main():
    """Main entry point"""
    import argparse

    parser = argparse.ArgumentParser(
        description='Email Clustering System for macOS Apple Mail'
    )
    parser.add_argument(
        '--database',
        help='Path to Excel database file',
        default=None
    )
    parser.add_argument(
        '--limit',
        type=int,
        help='Maximum number of emails to process',
        default=50
    )

    args = parser.parse_args()

    # Create clusterer instance
    clusterer = EmailClusterer(database_path=args.database)

    # Process emails
    clusterer.process_emails(limit=args.limit)


if __name__ == '__main__':
    main()
