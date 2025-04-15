import win32com.client
from datetime import datetime, timedelta, timezone

days_ahead = 8  # Number of days to look ahead for meetings

# Initialize Outlook COM object
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def meetings_ahead(namespace, days_ahead):
    # Time range
    now = datetime.now(timezone.utc)
    end_time = now + timedelta(days=days_ahead)

    # Get all accounts
    accounts = namespace.Accounts
    
    meetings_per_account = {}

    for account in accounts:
        try:
            print(f"\nAccount: {account.SmtpAddress}")
            
            # Access root folder and Calendar folder
            root_folder = namespace.Folders(account.DisplayName)
            calendar_folder = root_folder.Folders("Calendar")

            # Get calendar items and configure view
            items = calendar_folder.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True

            # Collect meetings
            meetings = []

            for item in items:
                try:
                    start = item.Start
                    end = item.End
                    
                    # Skip items without start time
                    if start is None:
                        continue

                    # Convert start to timezone-aware
                    if not hasattr(start, 'tzinfo') or start.tzinfo is None:
                        start = start.replace(tzinfo=timezone.utc)
                    if not hasattr(end, 'tzinfo') or end.tzinfo is None:
                        end = end.replace(tzinfo=timezone.utc)

                    # Filter by time range
                    if now <= start <= end_time:
                        meetings.append({
                            "subject": item.Subject,
                            "start": start,
                            "end": end
                        })
                        # print(f"Meeting: {item.Subject}, Start: {start}, End: {end}")
                except Exception as e:
                    print(f"Error processing item: {e}")

            meetings_per_account[account.SmtpAddress] = meetings

        except Exception as e:
            print(f"Error accessing account {account.DisplayName}: {e}")

    return meetings_per_account

# print(meetings_ahead(namespace, days_ahead))
meet = meetings_ahead(namespace, days_ahead)
for account, meetings in meet.items():
    print(f"\nAccount: {account}")
    for meeting in meetings:
        print(f"Subject: {meeting['subject']}, Start: {meeting['start']}, End: {meeting['end']}")
