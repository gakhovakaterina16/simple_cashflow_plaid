from plaid import Client
from datetime import datetime
import httplib2 
from apiclient import discovery
from oauth2client.service_account import ServiceAccountCredentials
import json
from config import CLIENT_ID, SECRET_KEY, PUBLIC_KEY, ENV, \
API_VERSION, CREDENTIALS_FILE, EMAIL

# Settings
# for Plaid
client = Client(
  client_id=CLIENT_ID,
  secret=SECRET_KEY,
  public_key=PUBLIC_KEY,
  environment=ENV,
  api_version=API_VERSION
)

# Generate a public_token for a given institution ID
# and set of initial products
# Code from Plaid Docs
create_response = \
    client.Sandbox.public_token.create('ins_1',
        ['transactions'])
# The generated public_token can now be
# exchanged for an access_token
exchange_response = \
    client.Item.public_token.exchange(
        create_response['public_token'])

access_token = exchange_response['access_token']

response_for_transactions = client.Transactions.get(access_token,
                                   start_date='2020-03-01',
                                   end_date='2020-05-20')
transactions = response_for_transactions['transactions']

# Manipulate the count and offset parameters to paginate
# transactions and retrieve all available data
# by default: start_date = 2020-03-01, end_date = 2020-05-20
# Code from Plaid Docs
while len(transactions) < response_for_transactions['total_transactions']:
    response = client.Transactions.get(access_token,
                                       start_date='2020-03-01',
                                       end_date='2020-05-20',
                                       offset=len(transactions)
                                      )
    transactions.extend(response['transactions'])

# Settings for Goole Sheets

credentials = ServiceAccountCredentials.from_json_keyfile_name(
                                                               CREDENTIALS_FILE, 
                                                               ['https://www.googleapis.com/auth/spreadsheets', 
                                                               'https://www.googleapis.com/auth/drive']
                                                               )

httpAuth = credentials.authorize(httplib2.Http())
service = discovery.build('sheets', 'v4', http = httpAuth)

# create Google Sheets 
spreadsheet = service.spreadsheets().create(body = {
    'properties': {'title': 'Test Task', 'locale': 'ru_RU'},
    'sheets': [{'properties': {'sheetType': 'GRID',
                               'sheetId': 0,
                               'title': 'raw data',
                               'gridProperties': {'rowCount': 100, 'columnCount': len(transactions)+1}}},
               {'properties': {'sheetType': 'GRID',
                                'sheetId': 1,
                                'title': 'simple cashflow',
                                'gridProperties': {'rowCount': 100, 'columnCount': 15}}}]
}).execute()

# ID of created document
spreadsheetId = spreadsheet['spreadsheetId']

# authorization settings
driveService = discovery.build('drive', 'v3', http = httpAuth)
access = driveService.permissions().create(
    fileId = spreadsheetId,
    body = {'type': 'user', 'role': 'writer', 'emailAddress': EMAIL},  
    fields = 'id'
).execute()

# add raw transactions data to Google Sheets
results = service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheetId, body = {
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "raw data!B2:D5",
         "majorDimension": "ROWS",
         "values": [
                    [json.dumps(transactions)]
                   ]}
    ]
}).execute()

# upgrade data
needed_keys = ['amount', 'category', 'date']
for transaction in transactions: 
# delete needless keys
    for key in list(transaction.keys()):
        if key not in needed_keys:
            del transaction[key]
    
# for 'date' key: %Y-%m-%d => %b %Y
    transaction['date'] = datetime.strptime(transaction['date'], '%Y-%m-%d')
    transaction['date'] = datetime.strftime(transaction['date'], '%b %Y')
    
# category => category 1 and category 2
    categories_list = transaction.pop('category')
    transaction['Category 1'] = categories_list[0]
    try:
        transaction['Category 2'] = categories_list[1]
    except IndexError:
        transaction['Category 2'] = 'null'

# add income/expense to data, change the sign of amount
    if transaction['amount'] > 0:
        transaction['Income/Expense'] = 'Expense'
        transaction['amount'] = -transaction['amount']
    elif transaction['amount'] < 0:
        transaction['Income/Expense'] = 'Income'
        transaction['amount'] = -transaction['amount']

# try to build simple cashflow
results = service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheetId, body = {
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "simple cashflow!B2:I35",
         "majorDimension": "ROWS",
         "values": [
                    ['Income/Expense', 'Category 1', 
                     'Category 2', 'Apr 2020', 'May 2020', 'Grand total']
                   ]}
    ]
}).execute()

print('https://docs.google.com/spreadsheets/d/' + spreadsheetId)
