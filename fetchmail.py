# -*- coding: utf-8 -*-
"""
Created on Fri Oct 23 22:15:54 2020

@author: Hp
"""

while True:

#     from __future__ import print_function
    import base64  
    import pandas as pd
    import pickle
    import os.path
    from googleapiclient.discovery import build
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request

    # If modifying these scopes, delete the file token.pickle.
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

    # def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
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

    service = build('gmail', 'v1', credentials=creds)


    label = 'INBOX'
    print('\n')
    results = service.users().messages().list(userId='me',labelIds=label).execute()
    messages = results.get('messages',[])


    From=[]
    To=[]
    DateTime= []
    Subject= []
    Body= []


    def email():

    #     message_count = int(input('How many emails do you want to extract? '))
        message_count = len(messages)

        if not messages:
            print('No message found.')
        else:



            for message in messages[:message_count]:
                msg = service.users().messages().get(userId='me', id=message['id'],format = 'full', metadataHeaders=None).execute()




                from1 = [msg['payload']['headers'][4]['value'].split(';')[0].strip()]
                to = [msg['payload']['headers'][5]['value'].split(';')[0].strip()]
                dateTime = [msg['payload']['headers'][1]['value'].split(';')[0].strip()]
                subject = [msg['payload']['headers'][3]['value'].split(';')[0].strip()]

                From.append(from1)
                To.append(to)
                DateTime.append(dateTime)
                Subject.append(subject)



                try: 
                    body_1 = base64.urlsafe_b64decode(msg['payload']['parts'][0]['body']['data'].encode('utf-8'))
                    body_1 = body_1.decode('utf-8')
                    Body.append(body_1)

                except:
                    body_2 = base64.urlsafe_b64decode(msg['payload']['parts'][0]['parts'][0]['body']['data'].encode('utf-8'))
                    body_2 = body_2.decode('utf-8')
                    Body.append(body_2)


                try:
                    attachment = base64.urlsafe_b64decode(msg['payload']['parts'][1]['body']['attachmentId'].encode('utf-8'))
                    attachment = attachment.decode('utf-8')
                except:
                    pass




                columns= ["From", "To", "DateTime", "Subject", "Body"]
                Data = pd.DataFrame(columns=columns)
                Data["From"] = From
                Data["To"] = To
                Data['DateTime']= DateTime
                Data["Subject"]= Subject
                Data['Body'] = Body
        #             print(Data)
                Data.to_csv(index=False, encoding='utf-8') 
                Data.to_csv('Data.csv', encoding='utf-8')
        #             print(Data.shape)


    email()
    pd.options.display.max_colwidth = 100000000
    df = pd.read_csv('Data.csv',encoding='utf-8')
    df.drop(['Unnamed: 0', 'To'], axis='columns', inplace=True)
    df = df.replace('\]','',regex=True)
    df = df.replace('\[','',regex=True)
    df = df.replace('\<','',regex=True)
    df = df.replace('\>','',regex=True)
    df = df.replace("\'",'',regex=True)
    df = df.replace("\\r\n",'',regex=True)
    df
    df.to_csv('final_Email_data.csv')


    print('Email Data has been successfully saved to the Directory')
