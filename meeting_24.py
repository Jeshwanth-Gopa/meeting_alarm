import win32com.client
from datetime import datetime, timedelta, timezone

days_ahead = 8  # Number of days to look ahead for meetings

# Initialize Outlook COM object
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Get all accounts
accounts = namespace.Accounts

# Time range
now = datetime.now(timezone.utc)
end_time = now + timedelta(days=days_ahead)

# Store meetings per account
meetings_per_account = {}

print("Fetching upcoming meetings in next " + str(days_ahead) + " days:\n")

for account in accounts:
    try:
        print(f"\nAccount: {account.DisplayName} ({account.SmtpAddress})")
        
        # Access root folder and Calendar folder
        root_folder = namespace.Folders(account.DisplayName)
        calendar_folder = root_folder.Folders("Calendar")

        # Get calendar items and configure view
        items = calendar_folder.Items
        # print(items)
        # items.Sort("[Start]")
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

            except Exception as e:
                print(f"Error reading item: {e}")

        meetings_per_account[account.SmtpAddress] = meetings

        # Print summary
        if meetings:
            print(f"✅ Found {len(meetings)} meeting(s).")
        else:
            print("ℹ️ No meetings in next " + str(days_ahead) + " days.")

    except Exception as e:
        print(f"⚠️ Could not access calendar for {account.DisplayName}: {e}")

# Display collected meetings
print("\nSummary of Meetings:")
for account, meetings in meetings_per_account.items():
    print(f"\nAccount: {account}")
    for meeting in meetings:
        print(f"- {meeting['subject']} | Start: {meeting['start']} | End: {meeting['end']}")