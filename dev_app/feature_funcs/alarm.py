import win32com.client
import heapq
from datetime import datetime, timedelta
from meetings_ahead import meetings_ahead, remove_timezone

days_ahead = 30  # Number of days to look ahead for meetings
ring_before = 5  # Minutes before the meeting to ring

outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

def ring_time(meetings):
    # for item 
    heap = [(item["ring_at"]+item["snoozeed"], item) for item in meetings]
    heapq.heapify(heap)
    _, meeting = heapq.heappop(heap)
    return meeting

def snooze(meeting, meetings):
    for meet in meetings:
        if meeting["id"] == meet["id"]:
            meetings.remove(meet)
            break
    meetings.append(meeting)
    return ring_time(meetings)

meeting = meetings_ahead(namespace, days_ahead, ring_before)
meet = ring_time(meeting)
print(f"Subject: {meet['subject']}, Start: {meet['start']}, End: {meet['end']}, Account: {meet['account']}, Ring at: {meet['ring_at']}")