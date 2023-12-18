import googleapiclient.discovery
import gdown
import pandas as pd
from datetime import datetime, timedelta
import pytz
from dateutil import parser
import os

# Function to authenticate google calendar
def authenticate_google_calendar(api_key):
    return googleapiclient.discovery.build('calendar', 'v3', developerKey=api_key)

# Function to download xlsx file from Google Drive
def download_xlsx(url, output):
    gdown.download(url, output, quiet=True, fuzzy=True)

# Function to read xlsx file as pandas DataFrame
def read_xlsx(file_path, sheet_name):
    return pd.read_excel(file_path, sheet_name=sheet_name, keep_default_na=False)

def generate_birthday_messages(birthdays):
    cur_datetime = datetime.now()
    cur_month = int(cur_datetime.month)
    cur_day = int(cur_datetime.day)
    cur_year = int(cur_datetime.year)

    birthdays["birth_year"] = pd.to_numeric(birthdays["birth_year"], errors="coerce")

    birthdays["age"] = cur_year - birthdays["birth_year"]
    todays_events = birthdays.loc[(birthdays['birth_month'] == cur_month) & (birthdays['birth_date'] == cur_day)]

    message_list = []

    for x in range(len(todays_events)):
        message_dict = {}
        if len(todays_events['custom_message'].iloc[x]) > 1:
            message_dict['summary'] = todays_events['custom_message'].iloc[x]
        elif todays_events['congrats_to'].iloc[x].upper() == "ALL":
            message_dict['summary'] = "Drie violen, twee trommels en een fluit " + \
                todays_events['name'].iloc[x] + \
                " die is jarig en ziet er lekker uit! Ei ei ei en we zijn zo blij want " + \
                todays_events['name'].iloc[x] + \
                " die is jarig en dat vieren wij!"
        else:
            message_dict['summary'] = "Gefeliciteerd " + \
                todays_events['congrats_to'].iloc[x] + \
                " met " + \
                todays_events['name'].iloc[x] + \
                "!"
        if todays_events['age'].iloc[x] >= 1:
            message_dict['summary'] += " " + str(int(todays_events['age'].iloc[x])) + " jaar alweer, wat vliegt de tijd"

        message_list.append(message_dict)

    return message_list

def get_today_events(calendar_service, calendar_id):
    today_start = datetime.utcnow().replace(hour=0, minute=0, second=0, microsecond=0)
    today_end = today_start + timedelta(days=1)

    result = calendar_service.events().list(
        calendarId=calendar_id,
        timeMin=today_start.isoformat() + 'Z',
        timeMax=today_end.isoformat() + 'Z',
        singleEvents=True,
        orderBy='startTime'
    ).execute()

    return result.get('items', [])

def get_modified_events(calendar_service, calendar_id):
    # Set the time range to look for modifications in the last 24 hours
    end_time = datetime.utcnow()
    start_time = end_time - timedelta(days=1)

    # Convert start_time and end_time to UTC
    start_time_utc = start_time.replace(tzinfo=pytz.UTC)
    end_time_utc = end_time.replace(tzinfo=pytz.UTC)

    # Execute the API request to get modified events
    result = calendar_service.events().list(
        calendarId=calendar_id,
        timeMin=start_time_utc.isoformat(),
        singleEvents=True,
        orderBy='updated'
    ).execute()

    # Extract the 'updated' field from each event and filter for the last 24 hours
    modified_events = [event for event in result.get('items', []) if
                   parser.isoparse(event['updated']) >= start_time_utc]


    return modified_events

def print_calendar_events(title, events):
    print(f"\n*{title}:*")
    for event in events:
        summary = event['summary']
        creator_email = event.get('creator', {}).get('email', 'N/A')
        date = event.get('start', {}).get('dateTime', 'N/A').split('T')[0]
        time = event.get('start', {}).get('dateTime', 'N/A').split('T')[1][:5]

        print(f"\n_{summary} ({date}, {time}, {creator_email})_")

def print_birthday_messages(title, messages):
    print(f"\n*{title}:*")
    for message in messages:
        print(f"\n_{message['summary']}_")

def write_calendar_events(file, title, events):
    file.write(f"\n*{title}:*\n")
    for event in events:
        summary = event['summary']
        creator_email = event.get('creator', {}).get('email', 'N/A')

        start = event.get('start', {})
        end = event.get('end', {})
        is_all_day = 'date' in start and 'date' in end

        if is_all_day:
            start_date = start.get('date', 'N/A')
            end_date = end.get('date', 'N/A')
            time_str = f"{start_date} t/m {end_date}"
        elif 'dateTime' in start:
            date = start.get('dateTime', 'N/A').split('T')[0]
            time = start.get('dateTime', 'N/A').split('T')[1][:5]
            time_str = f"{date}, {time}"
        else:
            time_str = 'N/A'

        file.write(f"\n_{summary} ({time_str}, {creator_email})_\n")

def write_birthday_messages(file, title, messages):
    file.write(f"*{title}:*")
    for message in messages:
        file.write(f"\n_{message['summary']}_")

def main():
    # Calendar prep
    calendar_api_key = os.environ.get('CALENDAR_API_KEY')

    try:
        calendar_service = authenticate_google_calendar(calendar_api_key)
        calendar_id = os.environ.get('CALENDAR_ID')

        # Birthday prep
        url = os.environ.get('BIRTHDAY_URL')
        output = 'birthdays.xlsx'

        try:
            download_xlsx(url, output)
            birthdays = read_xlsx('birthdays.xlsx', sheet_name=0)
            birthday_messages = generate_birthday_messages(birthdays)

            # Print the birthday messages if they exist
            with open("birthday_messages.txt", "w") as file:
                if birthday_messages:
                    write_birthday_messages(file, "Botolas' verjaardagen", birthday_messages)

        except Exception as birthday_error:
            print(f"Error processing birthdays: {str(birthday_error)}")

        # Collect the events
        try:
            today_events = get_today_events(calendar_service, calendar_id)
            modified_events = get_modified_events(calendar_service, calendar_id)

            # Print the events if they exist
            with open("today_events.txt", "w") as file:
                if today_events:
                    write_calendar_events(file, "Vandaag op het programma", today_events)

            with open("modified_events.txt", "w") as file:
                if modified_events:
                    write_calendar_events(file, "Toegevoegd/gewijzigd afgelopen 24 uur in de vriendenagenda", modified_events)

        except Exception as events_error:
            print(f"Error processing events: {str(events_error)}")

    except Exception as auth_error:
        print(f"Error authenticating Google Calendar API: {str(auth_error)}")

    print("All seems to be done")

if __name__ == "__main__":
    main()
