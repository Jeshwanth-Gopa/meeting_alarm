import win32com.client
from datetime import datetime, timedelta
import time
import tzlocal
import pytz

# Get local timezone
local_tz = tzlocal.get_localzone()

def get_upcoming_meetings():
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar_folder = namespace.GetDefaultFolder(9)  # Calendar folder

    now = datetime.now(local_tz)
    end_time = now + timedelta(days=1)

    calendar_items = calendar_folder.Items
    calendar_items.IncludeRecurrences = True
    calendar_items.Sort("[Start]")

    meetings = []
    for item in calendar_items:
        start = item.Start.replace(tzinfo=local_tz)
        end = item.End.replace(tzinfo=local_tz)

        if now <= start <= end_time:
            meetings.append({
                "subject": item.Subject,
                "start": start,
                "end": end
            })

    return meetings

def check_and_alert(meetings):
    now = datetime.now(local_tz)

    for meeting in meetings:
        start_time = meeting["start"]
        time_to_meeting = (start_time - now).total_seconds() / 60

        print(f"Checking meeting: {meeting['subject']} | Start: {start_time} | Now: {now} | Time to meeting: {time_to_meeting}")

        # Alert if within next 5 minutes
        if 0 <= time_to_meeting <= 5:
            print(f"ALERT: Meeting '{meeting['subject']}' is starting soon at {start_time}!")

if __name__ == "__main__":
    meetings = get_upcoming_meetings()
    print("Meetings fetched. Monitoring...")

    while True:
        check_and_alert(meetings)
        time.sleep(60)
