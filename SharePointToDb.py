import datetime
from time import sleep
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import pyodbc 
import pandas as pd


## TODO: Add dotenv support
## TODO: Convert script to more generic form

url = 'https://##ParentName.sharepoint.com/:x:/r/sites/ParentDomain' ## this url can completely change according to your sharepoint site
username = '##UserEmail' ## Must be a SharePoint user with admin privilegess
password = '##Password'
relative_url = '/sites/##ParentDomain/##SubFolderIfExist/##ExcelNameInSharePoint.xlsx' ## this relative url can also change according to your sharepoint site pls test it before use

ctx_auth = AuthenticationContext(url) ## login sharepoint with credentials
if ctx_auth.acquire_token_for_user(username, password):
  ctx = ClientContext(url, ctx_auth)
  web = ctx.web
  ctx.load(web)
  ctx.execute_query()
  print("Web title: {0}".format(web.properties['Title']))

else:
  print (ctx_auth.get_last_error())

filename = '##ExcelName.xlsx' ## this is the name of the excel file that will be downloaded from sharepoint
with open(filename, 'wb') as output_file:
    response = File.open_binary(ctx, relative_url)
    output_file.write(response.content)

sleep(3) ## wait for 3 seconds to make sure that the file is downloaded

conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=##ServerIp;'
                      'Database=##DbName;'
                      'Trusted_Connection=yes;')

log_conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=##ServerIp;'
                      'Database=##DbName;'
                      'Trusted_Connection=yes;')


## Cursor for main table and log table
cursor = conn.cursor()
log_cursor = log_conn.cursor()

## Before insert the data to table, truncate the table to prevent duplicate data
cursor.execute("TRUNCATE TABLE ##MainTableName")

try:
	df = pd.read_excel('##ExcelName.xlsx', sheet_name='##SheetName', skiprows=2) ## skipping first 2 rows because of header, this can change according to your excel file's headers
	df = df.fillna(0) ## fill empty cells with 0 to prevent error
	for index, row in df.iterrows(): ## these columns can be changed according to the excel file, please be careful with the data types, if you have a string column, you must use varchar in sql
		column1 = int(df.iloc[index][0])
		column2 = df.iloc[index][1]
		column3 = int(df.iloc[index][2])
		column4 = df.iloc[index][3]
		column5 = df.iloc[index][4]
		column6 = df.iloc[index][5]
		column7 = df.iloc[index][6]
		column8 = df.iloc[index][7]
		column9 = df.iloc[index][8]
		column10 = df.iloc[index][9]
		values = (column1, column2, column3, column4, column5, column6, column7, column8, column9, column10) ## assign variables to values to insert db
		
		sql = """INSERT INTO ##MainTableName (column1, column2, column3, column4, column5, column6, column7, column8, column9, column10) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""" ## prepare sql query to insert
		
		cursor.execute(sql, values) ## execute sql query with values
		cursor.commit()

	log_cursor.execute("INSERT INTO ##LoggedTableName (LogType, LogMessage, LogDate) VALUES (?,?,?)", ('Success',  "List updated successfully.", datetime.datetime.now())) ## insert success logs to db to check from grafana dashboard or other monitoring tools
	log_cursor.commit() 

except Exception as e:
	log_cursor.execute("INSERT INTO ##LoggedTableName (LogType, LogMessage, LogDate) VALUES (?,?,?)", ('Error',  str(e), datetime.datetime.now())) ## insert error logs to db to check from grafana dashboard or other monitoring tools
	log_cursor.commit()