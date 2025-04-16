import win32com.client
from datetime import datetime, timedelta
from alarm import ring_time
from meetings_ahead import meetings_ahead

days_ahead = 8  # Number of days to look ahead for meetings
ring_before = 5 # Minutes before the meeting to ring

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

meetings = meetings_ahead(namespace, days_ahead)
print(ring_time(meetings, ring_before))