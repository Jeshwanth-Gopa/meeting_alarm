# app.py
from flask import Flask, jsonify
from alarm.alarm import get_meeting_to_ring

app = Flask(__name__)

@app.route('/get_next_meeting', methods=['GET'])
def get_next_meeting():
    try:
        meeting = get_meeting_to_ring()
        if meeting:
            return jsonify({"status": "success", "meeting": meeting})
        else:
            return jsonify({"status": "success", "message": "No meetings found"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

if __name__ == '__main__':
    app.run(debug=True)