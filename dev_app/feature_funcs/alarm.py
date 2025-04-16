import win32com.client
from datetime import datetime, timedelta
from meetings_ahead import meetings_ahead, remove_timezone

days_ahead = 30  # Number of days to look ahead for meetings

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def ring_time(meetings, ring_before):
    next_meeting = datetime.now() + timedelta(days=days_ahead)
    dic = {}
    for acc,account_meetings in meetings.items():
        for meeting in account_meetings:
            start_time = remove_timezone(meeting["start"])
            if start_time < next_meeting:
                dic = {
                    "subject": meeting["subject"],
                    "start": start_time,
                    "end": remove_timezone(meeting["end"]),
                    "account": acc,
                    "ring_at": start_time - timedelta(minutes=ring_before)
                }
    if len(dic) > 0:
        return dic
    return None

# meetings = meetings_ahead(namespace, days_ahead)
# print(ring_time(meetings, 5))