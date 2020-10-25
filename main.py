from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from uuid import uuid4
import xlrd
from xlutils.copy import copy

FIRST_NAME_COL = 0
LAST_NAME_COL = 1
EMAIL_COL = 2
START_DATETIME_COL = 3
END_DATETIME_COL = 4
MEET_LINK_COL = 5


def authorize():
    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/calendar']
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds


def create_event(service, first_name, last_name, email, start_datetime, end_datetime):
    event = {"summary": first_name + " " + last_name,
             "start": {
                 "dateTime": start_datetime,
             },
             "end": {
                 "dateTime": end_datetime,
             },
             "attendees": [
                 {'email': email},
             ],
             "conferenceData": {"createRequest": {"requestId": f"{uuid4().hex}",
                                                  "conferenceSolutionKey": {"type": "hangoutsMeet"}}},
             "reminders": {"useDefault": True}
             }

    event = service.events().insert(calendarId='primary', body=event,
                                    conferenceDataVersion=1, sendUpdates='all').execute()
    print('Event created: %s' % (event.get('htmlLink')))

    meet_link = event.get('conferenceData')['entryPoints'][0]['uri']
    return meet_link


def main():

    rb = xlrd.open_workbook("excel/input.xls")  # open excel
    r_sheet = rb.sheet_by_index(0)  # read only copy to introspect the file
    wb = copy(rb)  # a writable copy (write only, no read)
    w_sheet = wb.get_sheet(0)  # the sheet to write to within the writable copy

    creds = authorize()
    service = build('calendar', 'v3', credentials=creds)

    # for each candidate
    for row in range(1, 4):
        try:
            first_name = r_sheet.cell_value(row, FIRST_NAME_COL)
        except IndexError:
            break
        last_name = r_sheet.cell_value(row, LAST_NAME_COL)
        email = r_sheet.cell_value(row, EMAIL_COL)
        start_datetime = r_sheet.cell_value(row, START_DATETIME_COL)
        end_datetime = r_sheet.cell_value(row, END_DATETIME_COL)

        # exit when there aren't any candidates left
        print(f"Row {row}. Name: {first_name} {last_name}")

        meet_link = create_event(service, first_name, last_name,
                                 email, start_datetime, end_datetime)

        w_sheet.write(row, MEET_LINK_COL, meet_link)
        wb.save('excel/output.xls')

        row += 1


if __name__ == '__main__':
    main()
