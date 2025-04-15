import win32com.client
from datetime import datetime, timedelta
import tzlocal

# Get the local timezone
local_tz = tzlocal.get_localzone()

def get_outlook_calendars():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendars = []
    
    # Iterate over top-level folders in Outlook.
    for folder in namespace.Folders:
        try:
            # Try accessing the Calendar folder within each account
            calendar_folder = folder.Folders["Calendar"]
            calendars.append({
                "account": folder.Name,  # or folder.Store.DisplayName if you prefer
                "calendar_folder": calendar_folder
            })
        except Exception as e:
            # Folder might not contain "Calendar" so we skip
            continue
    return calendars

def get_upcoming_meetings():
    """
    Retrieves meetings from all available account calendars (if present)
    within the next 8 days.
    """
    calendars = get_outlook_calendars()
    now = datetime.now(local_tz)
    end_time = now + timedelta(days=8)
    meetings = []

    for calendar in calendars:
        account_name = calendar["account"]
        calendar_folder = calendar["calendar_folder"]

        # Get all items from the calendar
        items = calendar_folder.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")

        for item in items:
            try:
                start = item.Start.replace(tzinfo=local_tz)
                end   = item.End.replace(tzinfo=local_tz)
                if now <= start <= end_time:
                    meetings.append({
                        "subject": item.Subject,
                        "start_time": start,
                        "end_time": end,
                        "account": account_name  # Save the account name here
                    })
            except Exception as e:
                print(f"Error processing item: {e}")

    return meetings
