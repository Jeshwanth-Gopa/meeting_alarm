import win32com.client
import pywintypes
from datetime import datetime, timedelta, timezone

days_ahead = 8  # Number of days to look ahead for meetings

# Initialize Outlook COM object
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def remove_timezone(dt):
    # Remove the timezone information, making it naive if it is aware
    if dt.tzinfo:
        return dt.replace(tzinfo=None)
    return dt

def meetings_ahead(namespace, days_ahead):
    """
    Parameters:
    namespace (object): The namespace object containing the accounts to fetch calendar data from.
    days_ahead (int): The number of days from the current date to filter meetings.

    Returns:
    dict: 
          Example:
            {
                "account1@example.com": [
                    {"subject": "Meeting 1", "start": datetime1, "end": datetime2},
                    {"subject": "Meeting 2", "start": datetime3, "end": datetime4}
                ],
                "account2@example.com": [
                    {"subject": "Meeting 3", "start": datetime5, "end": datetime6}
                ]
            }
    """
    # Time range
    now = datetime.now()
    end_time = now + timedelta(days=days_ahead)

    # Get all accounts
    accounts = namespace.Accounts
    
    meetings_per_account = {}

    for account in accounts:
        try:
            # print(f"\nAccount: {account.SmtpAddress}")
            
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
                    start = remove_timezone(item.Start)
                    end = remove_timezone(item.End)
                    
                    # Skip items without start time
                    if start is None:
                        continue

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

# meet = meetings_ahead(namespace, days_ahead)
# for account, meetings in meet.items():
#     print(f"\nAccount: {account}")
#     for meeting in meetings:
#         print(f"Subject: {meeting['subject']}, Start: {meeting['start']}, End: {meeting['end']}")