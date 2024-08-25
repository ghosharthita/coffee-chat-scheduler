import os
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build
from datetime import datetime, timedelta

# Initialize the Slack app
app = App(token=os.environ["SLACK_BOT_TOKEN"])

# Google Calendar API setup
SCOPES = ['https://www.googleapis.com/auth/calendar']

def get_calendar_service(user_id):
    # Load credentials from the environment variable
    creds = service_account.Credentials.from_service_account_file(
        f'token_{user_id}.json', SCOPES)
    creds = Credentials.from_authorized_user_file(f'token_{user_id}.json', SCOPES)
    return build('calendar', 'v3', credentials=creds)

def get_free_slots(service, attendees, time_min, time_max):
    free_busy_query = {
        "timeMin": time_min,
        "timeMax": time_max,
        "items": [{"id": attendee} for attendee in attendees]
    }
    
    events_result = service.freebusy().query(body=free_busy_query).execute()
    calendars = events_result.get('calendars', {})
    
    # Combine all busy periods
    all_busy = []
    for calendar in calendars.values():
        all_busy.extend(calendar.get('busy', []))
    
    # Sort busy periods
    all_busy.sort(key=lambda x: x['start'])
    
    # Find free slots
    free_slots = []
    current_time = datetime.fromisoformat(time_min[:-1])
    for busy in all_busy:
        busy_start = datetime.fromisoformat(busy['start'][:-1])
        if current_time < busy_start:
            free_slots.append((current_time, busy_start))
        current_time = max(current_time, datetime.fromisoformat(busy['end'][:-1]))
    
    if current_time < datetime.fromisoformat(time_max[:-1]):
        free_slots.append((current_time, datetime.fromisoformat(time_max[:-1])))
    
    return free_slots[:3]  # Return top 3 free slots

@app.command("/reschedule")
def reschedule_command(ack, say, command):
    ack()
    user_id = command['user_id']
    service = get_calendar_service(user_id)
    
    # Get the event ID from the command text
    # Assuming the command is used like: /reschedule event_id
    event_id = command['text'].strip()
    
    try:
        event = service.events().get(calendarId='primary', eventId=event_id).execute()
    except Exception as e:
        say(f"Error: Unable to find the specified event. Please check the event ID and try again.")
        return
    
    attendees = [attendee['email'] for attendee in event.get('attendees', [])]
    attendees.append('primary')  # Include the user's primary calendar
    
    now = datetime.utcnow().isoformat() + 'Z'
    end = (datetime.utcnow() + timedelta(days=14)).isoformat() + 'Z'  # Look for slots in the next 14 days
    
    free_slots = get_free_slots(service, attendees, now, end)
    
    options = []
    for i, (start, end) in enumerate(free_slots):
        options.append({
            "text": {
                "type": "plain_text",
                "text": f"{start.strftime('%Y-%m-%d %H:%M')} - {end.strftime('%Y-%m-%d %H:%M')}"
            },
            "value": f"slot_{i}_{event_id}"
        })
    
    blocks = [
        {
            "type": "section",
            "text": {
                "type": "mrkdwn",
                "text": f"Here are 3 available slots for rescheduling '{event.get('summary', 'Unnamed event')}':"
            }
        },
        {
            "type": "actions",
            "elements": [
                {
                    "type": "static_select",
                    "placeholder": {
                        "type": "plain_text",
                        "text": "Select a time slot"
                    },
                    "options": options,
                    "action_id": "slot_selected"
                }
            ]
        }
    ]
    
    say(blocks=blocks)

@app.action("slot_selected")
def handle_slot_selection(ack, body, say):
    ack()
    selected_slot = body['actions'][0]['selected_option']['value']
    slot_index, event_id = selected_slot.split('_')[1:]
    slot_index = int(slot_index)
    
    user_id = body['user']['id']
    service = get_calendar_service(user_id)
    
    event = service.events().get(calendarId='primary', eventId=event_id).execute()
    attendees = [attendee['email'] for attendee in event.get('attendees', [])]
    attendees.append('primary')
    
    now = datetime.utcnow().isoformat() + 'Z'
    end = (datetime.utcnow() + timedelta(days=14)).isoformat() + 'Z'
    
    free_slots = get_free_slots(service, attendees, now, end)
    selected_start, selected_end = free_slots[slot_index]
    
    # Update the event with the new time
    event['start'] = {'dateTime': selected_start.isoformat(), 'timeZone': 'UTC'}
    event['end'] = {'dateTime': selected_end.isoformat(), 'timeZone': 'UTC'}
    
    updated_event = service.events().update(calendarId='primary', eventId=event_id, body=event).execute()
    
    say(f"Great! I've rescheduled '{updated_event.get('summary', 'Unnamed event')}' for {selected_start.strftime('%Y-%m-%d %H:%M')} - {selected_end.strftime('%Y-%m-%d %H:%M')}.")

if __name__ == "__main__":
    handler = SocketModeHandler(app, os.environ["SLACK_APP_TOKEN"])
    handler.start()