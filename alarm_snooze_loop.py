import time
from datetime import datetime, timedelta, timezone
import winsound  # For sound notification

# Function to check and alert for upcoming meetings
def check_and_alert(meetings_per_account):
    print('Called check_and_alert, meetings_per_account:', meetings_per_account)
    alert_threshold = timedelta(minutes=5)  # Alert 5 minutes before meeting
    
    while True:
        now = datetime.now(timezone.utc)

        for account, meetings in meetings_per_account.items():
            for meeting in meetings:
                start_time = meeting['start']
                time_to_meeting = start_time - now

                if timedelta(0) <= time_to_meeting <= alert_threshold:
                    print(f"\n🚨 Alert: Upcoming meeting for {account}!")
                    print(f"Subject: {meeting['subject']}")
                    print(f"Starts at: {start_time.strftime('%Y-%m-%d %H:%M:%S %Z')}\n")
                    
                    # Optional: Play a sound notification
                    duration = 1000  # milliseconds
                    freq = 1000  # Hz
                    winsound.Beep(freq, duration)

        time.sleep(60)  # Check every 60 seconds
