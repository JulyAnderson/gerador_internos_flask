from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

#ID das planilhas de trabalho
ALTERACAO = '1Yuygb6lu4j65KBxw35vzFFb52u_R2rORt0rcdVFTJR0'
HOMOLOGADO = '1G-Vu_mDa775qWCKcPnw6ys3XPP9DO1k1q6Cljhba3hY'

# The ID and range of a sample spreadsheet.
SAMPLE_RANGE_NAME = 'A:Z'

def get_planilhas_google(SAMPLE_SPREADSHEET_ID):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
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

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
            
    return values

# Obtendo os dados de ALTERACAO e HOMOLOGADO
BASE_ALTERACAO = get_planilhas_google(ALTERACAO)
BASE_HOMOLOGADO = get_planilhas_google(HOMOLOGADO)

