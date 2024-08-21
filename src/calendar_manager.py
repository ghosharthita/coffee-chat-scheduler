import msal
import requests
import datetime
import json
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.parse

# Azure AD app details
CLIENT_ID = "xxxx"
CLIENT_SECRET = "xxxx"
TENANT_ID = "xxxx"
REDIRECT_URI = "http://localhost:8000"

# Microsoft Graph API endpoint
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0'

# Authentication scopes
SCOPES = ['Calendars.Read']

auth_code = None

class AuthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        global auth_code
        query_components = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        auth_code = query_components.get('code', [None])[0]
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        self.wfile.write(b'Authentication successful! You can close this window now.')

def get_access_token():
    global auth_code
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )

    # Generate the authorization URL
    auth_url = app.get_authorization_request_url(
        SCOPES,
        redirect_uri=REDIRECT_URI
    )

    print(f"Please go to this URL to sign in: {auth_url}")
    webbrowser.open(auth_url)

    # Start local server to receive the authorization code
    httpd = HTTPServer(('localhost', 8000), AuthHandler)
    httpd.handle_request()

    if auth_code:
        result = app.acquire_token_by_authorization_code(
            auth_code,
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI
        )

        if "access_token" in result:
            print("Token acquired successfully")
            return result['access_token']
        else:
            print(f"Error: {result.get('error')}")
            print(f"Error description: {result.get('error_description')}")
            return None
    else:
        print("No authorization code received")
        return None

def get_calendar_events(access_token):
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Accept': 'application/json'
    }

    now = datetime.datetime.utcnow()
    end_date = now.isoformat() + 'Z'
    start_date = (now - datetime.timedelta(days=7)).isoformat() + 'Z'

    query_params = {
        'startDateTime': start_date,
        'endDateTime': end_date,
        'select': 'subject,start,end,organizer',
        'orderby': 'start/dateTime desc',  # Order by start time, most recent first
        '$top': 100  # Increase the number of events to fetch to ensure we get all from the past week
    }

    url = f'{GRAPH_API_ENDPOINT}/me/calendar/events'
    print(f"Requesting URL: {url}")
    print(f"Fetching events from {start_date} to {end_date}")
    
    response = requests.get(url, headers=headers, params=query_params)

    print(f"Response status code: {response.status_code}")
    print(f"Response headers: {json.dumps(dict(response.headers), indent=2)}")
    print(f"Response content: {response.text[:200]}...")  # Print first 200 characters of response

    if response.status_code == 200:
        events = response.json()['value']
        # Sort events by start time, most recent first
        events.sort(key=lambda x: x['start']['dateTime'], reverse=True)
        return events
    else:
        print(f"Error: {response.status_code} - {response.text}")
        return None

def main():
    access_token = get_access_token()
    if access_token:
        events = get_calendar_events(access_token)
        if events:
            print("\nCalendar events from the last 7 days:")
            for event in events:
                start = datetime.datetime.fromisoformat(event['start']['dateTime'][:-1]).strftime("%Y-%m-%d %H:%M")
                end = datetime.datetime.fromisoformat(event['end']['dateTime'][:-1]).strftime("%H:%M")
                subject = event['subject']
                organizer = event['organizer']['emailAddress']['name']
                print(f"{start} - {end}: {subject} (Organizer: {organizer})")
        else:
            print("No events found in the last 7 days or an error occurred.")
    else:
        print("Failed to acquire access token.")

if __name__ == '__main__':
    main()