
#While solving this problem, I encountered some Rate Limit errors caused by the limit of requests per minute allowed by Sheets API.  Because of that, if loaded more than one time, the app may be a bit slow.
#To minimize these issues, an approach was implemented using exponential backoff, as suggested in this article (by Google): https://developers.google.com/sheets/api/limits?hl=pt-br#exponential.
#This is the link for the google sheet: https://docs.google.com/spreadsheets/d/1Cac_PBcvPAuBZ2nGQalMsIV4beGf__tzmfYxH2iZsI4/edit#gid=0
#It was considered that the student with 15 (or more) classes missed has automatically failed. 

import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
import time

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1Cac_PBcvPAuBZ2nGQalMsIV4beGf__tzmfYxH2iZsI4"

MAX_RETRIES = 3
BASE_DELAY = 1  #Base delay in seconds
BACKOFF_FACTOR = 2  # Back-off factor

def main(): 
    credentials = None
    #Verifying if "Sheets API" credentials are already set.
    if os.path.exists("token.json"):
        credentials = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            credentials = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(credentials.to_json())
    try:
        service = build("sheets", "v4", credentials=credentials)
        sheets = service.spreadsheets()
        
        #Values are read between rows 4 and 28, and columns C and H
        for row in range(4, 28):
            #To avoid errors in rate limit, there are some control variables that define the scope of the "while" loop.
            retries = 0
            delay = BASE_DELAY
            while retries < MAX_RETRIES:
                try:
                    abs_value = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!C{row}").execute().get("values")[0][0])
                    if abs_value >= 15:
                        situation = "Reprovado por falta"
                    else:
                        #Here all the values are read and medium "m" and "naf" are calculated based on those values. After that, the situation is updated.
                        p1 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!D{row}").execute().get("values")[0][0])
                        p2 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!E{row}").execute().get("values")[0][0])
                        p3 = int(sheets.values().get(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!F{row}").execute().get("values")[0][0])
                        m = (p1 + p2 + p3) / 30 
                        print(f"Processing {p1} + {p2} + {p3} / 3")
                        if m < 5:
                            situation = "Reprovado por nota"
                        elif 5 <= m < 7:
                            situation = "Exame final"
                            naf = (10 - m)
                            sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!H{row}", valueInputOption="USER_ENTERED", body={"values": [[f"{naf}"]]}).execute()
                        else:
                            situation = "Aprovado"

                    sheets.values().update(spreadsheetId=SPREADSHEET_ID, range=f"engenharia_de_software!G{row}", valueInputOption="USER_ENTERED", body={"values": [[f"{situation}"]]}).execute()
                    
                    break  #If the update is well succeedeed, exit the loop
                except HttpError as error:
                    if error.resp.status == 429:  #An rate limit error is thrown. Exponential backoff is implemented to deal with this exception.
                        print(f"Rate limit exceeded. Retrying in {delay} seconds...")
                        retries += 1
                        time.sleep(delay)
                        delay *= BACKOFF_FACTOR  #Increases delay exponencially
                    else:
                        print(error)
                        break  #Leave the loop if an exception different than rate limit is thrown.
    except Exception as e:
        print(f"Unexpected error: {e}")

if __name__ == "__main__":
    main()
