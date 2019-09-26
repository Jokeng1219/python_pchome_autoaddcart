from __future__ import print_function
import time
import pickle
import os.path
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from selenium.webdriver.support.ui import WebDriverWait

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '' # spreadsheetId
SAMPLE_RANGE_NAME = 'sheet1!A2:E' # Your range

driver = webdriver.Chrome()

def main():

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
                'client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    # there world not be an error, since code works fine.
    
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        for row in values:
            cart(row)

def cart(row):
    driver.get(row[2])
    time.sleep(0.5)

    try:
        select = Select(driver.find_element_by_xpath("//*[@id='ButtonContainer']/select"))
        select.select_by_value(row[3])
    except:
        pass

    element = driver.find_element_by_xpath('//*[@id="ButtonContainer"]/button')
    driver.execute_script("arguments[0].click();", element)

    try:
        WebDriverWait(driver, 1).until(EC.alert_is_present())

        alert = driver.switch_to.alert
        alert.accept()
        print(row[1],'Sold Out')
        
    except:
        print(row[1],'數量：',row[3],'Add Cart Completely')

if __name__ == '__main__':
    main()

    print('--Done--')