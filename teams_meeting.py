import win32com.client
from datetime import datetime, timedelta
import time
import tzlocal
from playsound import playsound

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
        try:
            start = item.Start.replace(tzinfo=local_tz)
            end = item.End.replace(tzinfo=local_tz)

            if now <= start <= end_time:
                meetings.append({
                    "subject": item.Subject,
                    "start": start,
                    "end": end
                })
        except Exception as e:
            print(f"Error processing item: {e}")

    return meetings

def check_and_alert(meetings):
    now = datetime.now(local_tz)

    for meeting in meetings:
        start_time = meeting["start"]
        time_to_meeting = (start_time - now).total_seconds() / 60

        print(f"Checking meeting: {meeting['subject']} | Start: {start_time} | Now: {now} | Time to meeting: {time_to_meeting:.2f} mins")

        # Alert if within next 5 minutes
        if 0 <= time_to_meeting <= 5:
            print(f"ALERT: Meeting '{meeting['subject']}' is starting soon at {start_time}!")
            playsound("ring1.wav")

if __name__ == "__main__":
    print("Meeting alarm started. Monitoring your Outlook calendar...")

    while True:
        try:
            meetings = get_upcoming_meetings()
            check_and_alert(meetings)
            time.sleep(60)  # Check every 60 seconds
        except KeyboardInterrupt:
            print("Program interrupted by user. Exiting...")
            break
        except Exception as e:
            print(f"An error occurred: {e}")
            time.sleep(60)