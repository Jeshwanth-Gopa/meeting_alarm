# app.py
import os
import json
from datetime import datetime, timedelta
import pytz
from tzlocal import get_localzone
from flask import Flask, render_template, redirect, url_for, flash
from apscheduler.schedulers.background import BackgroundScheduler
from win10toast import ToastNotifier
from outlook_fetcher import get_upcoming_meetings  # Import our fetcher

# ----- Configuration & Setup -----
app = Flask(__name__)
app.secret_key = "secret-key-for-session"

scheduler = BackgroundScheduler()
scheduler.start()

notifier = ToastNotifier()

local_tz = get_localzone()
JSON_FILE = "meetings.json"
MEETING_WINDOW_DAYS = 8

# ----- JSON Persistence Functions -----
def load_meetings():
    if os.path.exists(JSON_FILE):
        with open(JSON_FILE, "r") as file:
            try:
                data = json.load(file)
                # Convert ISO strings back to datetime objects
                for meeting in data:
                    meeting["start_time"] = datetime.fromisoformat(meeting["start_time"]).astimezone(local_tz)
                return data
            except json.JSONDecodeError:
                return []
    return []

def save_meetings(meetings):
    serializable_meetings = []
    for m in meetings:
        serializable_meetings.append({
            "id": m["id"],
            "subject": m["subject"],
            "start_time": m["start_time"].isoformat(),
            "alert_job_id": m.get("alert_job_id", None)
        })
    with open(JSON_FILE, "w") as file:
        json.dump(serializable_meetings, file)

def get_next_meeting_id(meetings):
    if not meetings:
        return 1
    return max(m["id"] for m in meetings) + 1

# ----- Notification Alert Function -----
def alert_meeting(meeting_id, subject, start_time):
    msg = f"Meeting '{subject}' starting at {start_time.strftime('%H:%M')}"
    print("ALERT:", msg)
    notifier.show_toast("Meeting Alert", msg, duration=10, threaded=True)
    # Remove the meeting after the alert
    meetings = load_meetings()
    meetings = [m for m in meetings if m["id"] != meeting_id]
    save_meetings(meetings)

# ----- Outlook Meetings Synchronization -----
def update_meetings_from_outlook():
    """Fetch meetings from Outlook and integrate them into the local JSON store."""
    outlook_meetings = get_upcoming_meetings()
    meetings = load_meetings()
    now = datetime.now(local_tz)

    # For simplicity, use the subject and start_time to determine uniqueness
    existing = {(m["subject"], m["start_time"].isoformat()) for m in meetings}

    updated = False
    for om in outlook_meetings:
        key = (om["subject"], om["start_time"].isoformat())
        # Only add if meeting not already present and meeting is in the future
        if key not in existing and om["start_time"] > now:
            meeting_id = get_next_meeting_id(meetings)
            # Calculate alert time (5 minutes before the meeting)
            alert_time = om["start_time"] - timedelta(minutes=5)
            # Schedule the alert only if alert time is in the future
            if alert_time > now:
                job = scheduler.add_job(
                    alert_meeting,
                    'date',
                    run_date=alert_time,
                    args=[meeting_id, om["subject"], om["start_time"]]
                )
                new_meeting = {
                    "id": meeting_id,
                    "subject": om["subject"],
                    "start_time": om["start_time"],
                    "alert_job_id": job.id,
                    "account": om.get("account", "Unknown")  # Add the account info
                }
                meetings.append(new_meeting)
                updated = True
                print(f"Added meeting: {om['subject']} at {om['start_time']} from account: {new_meeting['account']}")
    if updated:
        save_meetings(meetings)
    else:
        print("No new meetings to add from Outlook.")


# Schedule this function to run periodically (e.g., every 10 minutes)
scheduler.add_job(update_meetings_from_outlook, 'interval', minutes=10)

# Optionally, call it at startup so it runs immediately.
update_meetings_from_outlook()

# ----- Flask Route (GET-only, because no manual new meeting form) -----
@app.route("/", methods=["GET"])
def index():
    meetings = load_meetings()
    now = datetime.now(local_tz)
    window_end = now + timedelta(days=MEETING_WINDOW_DAYS)
    # Display only meetings within the 8-day window.
    meetings = [m for m in meetings if now <= m["start_time"] <= window_end]
    return render_template("index.html", meetings=meetings)

# Routes for cancel/snooze remain as in your current code...
@app.route("/cancel/<int:meeting_id>")
def cancel_meeting(meeting_id):
    meetings = load_meetings()
    meeting = next((m for m in meetings if m["id"] == meeting_id), None)
    if meeting:
        if meeting.get("alert_job_id"):
            try:
                scheduler.remove_job(meeting["alert_job_id"])
            except Exception as e:
                print("Error removing job:", e)
        meetings = [m for m in meetings if m["id"] != meeting_id]
        save_meetings(meetings)
        flash("Meeting alert cancelled.", "success")
    else:
        flash("Meeting not found.", "error")
    return redirect(url_for("index"))

@app.route("/snooze/<int:meeting_id>")
def snooze_meeting(meeting_id):
    meetings = load_meetings()
    meeting = next((m for m in meetings if m["id"] == meeting_id), None)
    if meeting:
        if meeting.get("alert_job_id"):
            try:
                scheduler.remove_job(meeting["alert_job_id"])
            except Exception as e:
                print("Error removing job:", e)
        new_alert_time = datetime.now(local_tz) + timedelta(minutes=10)
        # Ensure alert does not exceed meeting start time
        if new_alert_time > meeting["start_time"]:
            new_alert_time = meeting["start_time"] - timedelta(seconds=10)
        job = scheduler.add_job(
            alert_meeting,
            'date',
            run_date=new_alert_time,
            args=[meeting_id, meeting["subject"], meeting["start_time"]]
        )
        meeting["alert_job_id"] = job.id
        save_meetings(meetings)
        flash(f"Meeting '{meeting['subject']}' snoozed until {new_alert_time.strftime('%H:%M')}", "success")
    else:
        flash("Meeting not found.", "error")
    return redirect(url_for("index"))

if __name__ == "__main__":
    app.run(debug=True)