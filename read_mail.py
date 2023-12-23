from simplegmail import Gmail
from simplegmail.query import construct_query
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from datetime import datetime, timedelta
import os.path
import pandas as pd
import re

SCOPES = ['https://www.googleapis.com/auth/calendar']

gmail = Gmail()

query_params = {
    "newer_than": (2, "day"),
    "sender": "noreply@cinema-city.pl",
}

creds = None
if os.path.exists("token.json"):
   creds = Credentials.from_authorized_user_file("token.json")
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
      creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token:
      token.write(creds.to_json())

service = build('calendar', 'v3', credentials=creds)
messages = gmail.get_messages(query=construct_query(query_params))

for message in messages:
    tables = pd.read_html(message.html, match='Order details')
    print(len(tables))
    result = tables[0][0][11]
    date_regex = '(\\d{2})/(\\d{2})/(\\d{4})\\s(\\d{2}):(\\d{2})'
    title_regex = '(?<=(\\d{2})/(\\d{2})/(\\d{4})\\s(\\d{2}):(\\d{2}))(.*)(?=\\sSala)'
    try:
        date_string_touple = re.search(date_regex, result).groups()
        date_int_touple = tuple(map(int, date_string_touple))
        touple_list = list(date_int_touple)
        touple_list[0], touple_list[2] = touple_list[2], touple_list[0]
        showing_date = datetime(*touple_list[0:5])
        showing_title = re.search(title_regex, result).groups()[5]
        start = showing_date.isoformat()
        end = (showing_date + timedelta(hours=3)).isoformat()
        event = {
            'summary': f'Kino: {showing_title}',
            'colorId': 2,
            'start': {
              'dateTime': start,
              'timeZone': 'Europe/Vienna',
            },
            'end': {
              'dateTime': end,
              'timeZone': 'Europe/Vienna',
            },
        }                 
        event = service.events().insert(calendarId='primary', body=event).execute()
        print('Event created: %s' % event.get('htmlLink'))
        print(showing_date)
        print(showing_title)
    except HttpError as error:
        print("Error occured", error)
    except AttributeError as error:
        print(error)
    except ValueError as error:
        print(error)
