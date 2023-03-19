import smtplib, ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
import yfinance as yf
import datetime
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import plotly.express as px
import os

scopes = [
'https://www.googleapis.com/auth/spreadsheets',
'https://www.googleapis.com/auth/drive'
]

# reading excel
credentials = ServiceAccountCredentials.from_json_keyfile_name("stockdata-381117-455fbd620cbc.json", scopes) #access the json key you downloaded earlier 
file = gspread.authorize(credentials) # authenticate the JSON key with gspread
sheet = file.open("stock_data") #open sheet
sheet = sheet.sheet1

data = pd.DataFrame(sheet.get_all_values(),columns=['stock_symbol', 'buy','sell'])
excel_stocks = data.stock_symbol.to_list()
length = len(data.stock_symbol)

# print(length)

trick = ["DLF.NS","ZYDUSLIFE.NS"]
# "NIFTYBEES.NS",

df = pd.DataFrame(columns=['stock','Action'])

for i in trick:
    try:
        ticker = yf.Ticker(i)
        # Retrieve today's data
        today_data = ticker.history(period="1d")

        open = int(today_data.Open.values)
        close = int(today_data.Close.values)

        if i in excel_stocks:
            print(i,"is there")
            for j in range(length):
                n = ( data[data['stock_symbol']==data.stock_symbol[j]].index.values[0])
                buy = int(data['buy'][n])
                sell = int(data['sell'][n])
                print(buy," ",sell)

                if (close<=buy):
                    print('buy')
                    new_data = {'stock': i, 'Action': 'buy'}
                    df = df.append(new_data, ignore_index=True)
                    print(df)
                elif (close>=sell):
                    print('sell')
                    new_data = {'stock': i, 'Action': 'sell'}
                    df = df.append(new_data, ignore_index=True)
                    print(df)
                else:
                    print('donothing')
                    new_data = {'stock': i, 'Action': 'donothing'}
                    df = df.append(new_data, ignore_index=True)

                print(n)
        else:
            print(i,"not there")
    except:
        print("error")











# Convert the DataFrame to an HTML table string
html_table = df.to_html(index=False)

css_styles = '''
<style>
table {
    border-collapse: collapse;
    width: 100%;
}

th, td {
    text-align: left;
    padding: 8px;
}

th {
    background-color: #004d99;
    color: white;
}
</style>
'''

# Define the HTML document
# Add an image element
##############################################################
html = '''
    <html>
    <head>{css}</head>
        <body>
            <h1>Action from NeBULA</h1>
            <p>Do Action ASAP!</p>
            {html_table}

        </body>
    </html>
    '''.format(html_table=html_table,css = css_styles)
##############################################################

   

# Set up the email addresses and password. Please replace below with your email address and password

email_from = 'bot.nebula9@gmail.com'
password = os.environ['PASSWORD']
email_to = 'naveenpraba08@gmail.com'
# Generate today's date to be included in the email Subject
date_str = pd.Timestamp.today().strftime('%Y-%m-%d')

# Create a MIMEMultipart class, and set up the From, To, Subject fields
email_message = MIMEMultipart()
email_message['From'] = email_from
email_message['To'] = email_to
email_message['Subject'] = f'Action_Report_from_nebula_bot_v4'

# Attach the html doc defined earlier, as a MIMEText html content type to the MIME message
email_message.attach(MIMEText(html, "html"))


email_string = email_message.as_string()
context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(email_from, password)
    server.sendmail(email_from, email_to, email_string) 
