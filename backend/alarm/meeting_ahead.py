import win32com.client
import hashlib
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

def meetings_ahead(namespace, days_ahead, ring_before=0):
    # Time range
    now = datetime.now()
    end_time = now + timedelta(days=days_ahead)

    # Get all accounts
    accounts = namespace.Accounts
    
    meetings = []

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

            for item in items:
                try:
                    start = remove_timezone(item.Start)
                    end = remove_timezone(item.End)
                    
                    # Skip items without start time
                    if start is None:
                        continue

                    # Filter by time range
                    if now <= start <= end_time:
                        st = item.Subject + str(start) + str(end) + account.SmtpAddress
                        meetings.append({
                            "subject": item.Subject,
                            "start": start,
                            "end": end,
                            "account": account.SmtpAddress,
                            "ring_at": start - timedelta(minutes=ring_before),
                            "snoozed": 0,
                            "id": hashlib.sha256(st.encode()).hexdigest()
                        })
                        # print(f"Meeting: {item.Subject}, Start: {start}, End: {end}")
                except Exception as e:
                    print(f"Error processing item: {e}")


        except Exception as e:
            print(f"Error accessing account {account.DisplayName}: {e}")

    return meetings

meet = meetings_ahead(namespace, days_ahead)
# print(meet)
# for meetings in meet:
#     print(f"Subject: {meetings['subject']}, Start: {meetings['start']}, End: {meetings['end']}, Account: {meetings['account']}, Ring at: {meetings['ring_at']}")
# print(meet)