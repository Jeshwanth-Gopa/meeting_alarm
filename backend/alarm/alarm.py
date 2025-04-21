# alarm/alarm.py
import win32com.client
import heapq
from meetings_ahead import meetings_ahead

def get_meeting_to_ring(days_ahead=30, ring_before=5):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    meetings = meetings_ahead(namespace, days_ahead, ring_before)
    if not meetings:
        return None
    heap = [(item["ring_at"] + item["snoozed"], item) for item in meetings]
    heapq.heapify(heap)
    _, meeting = heapq.heappop(heap)
    return meeting