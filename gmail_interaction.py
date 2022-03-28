from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

"""Send an email message from the user's account.
"""
##pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib

import base64
import email
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
import os
import docx2txt

from apiclient import errors

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']


##########################################
############ READING EMAILS ##############
##########################################

def get_messages(service, user_id):
  try:
    return service.users().messages().list(userId=user_id).execute()
  except Exception as error:
    print('An error occurred: %s' % error)


def get_message(service, user_id, msg_id):
  try:
    return service.users().messages().get(userId=user_id, id=msg_id, format='metadata').execute()
    print('Message snippet: %s' % message['snippet'])
  except Exception as error:
    print('An error occurred: %s' % error)  


def get_mime_message(service, user_id, msg_id):
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id, format='raw').execute()
    print('Message snippet: %s' % message['snippet'])
    msg_str = base64.urlsafe_b64decode(message['raw'].encode("utf-8")).decode("utf-8")
    mime_msg = email.message_from_string(msg_str)

    return mime_msg
  except Exception as error:
    print('An error occurred: %s' % error)

def get_attachments(service, user_id, msg_id, store_dir):
  try:
    message = service.users().messages().get(userId=user_id, id=msg_id).execute()

    for part in message['payload']['parts']:
      if(part['filename'] and part['body'] and part['body']['attachmentId']):
        attachment = service.users().messages().attachments().get(id=part['body']['attachmentId'], userId=user_id, messageId=msg_id).execute()

        file_data = base64.urlsafe_b64decode(attachment['data'].encode('utf-8'))
        path = ''.join([store_dir, part['filename']])

        f = open(path, 'wb')
        f.write(file_data)
        f.close()
  except Exception as error:
    print('An error occurred: %s' % error)

def read_email(service, user_id, my_subject):

  email_id = ""
  found = False
  page = get_messages(service, "me")



  email_list = page.get('messages')
  i = 0

  while (not found) and (i < len(email_list)):
    mail = email_list[i]
    potential_id = mail.get('id')

    ## get_message returns a mess of embedded dictionaries
    ## So this is us going through them to get to the subject line, which I think is the best search term
    ## dictionary -> dictionary -> list -> dictionary

    info = get_message(service, user_id, potential_id)
    payload = info.get('payload')
    header = payload.get('headers')
    
    for feature in header:
      function = feature.get('name')

      if 'Subject' in function:
        email_subject = feature.get('value')

        if my_subject in email_subject:
          found = True
          email_id = potential_id

    i+= 1

  print(email_id)

def parse_msg(msg):
    if msg.get("payload").get("body").get("data"):
        return base64.urlsafe_b64decode(msg.get("payload").get("body").get("data").encode("ASCII")).decode("utf-8")
    return msg.get("snippet") 
  
    

##########################################
########### DRAFTING EMAILS ##############
##########################################

def create_draft(service, user_id, message_body):
  """Create and insert a draft email. Print the returned draft's message and id.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message_body: The body of the email message, including headers.

  Returns:
    Draft object, including draft id and message meta data.
  """
  try:
    message = {'message': message_body}
    draft = service.users().drafts().create(userId=user_id, body=message).execute()

    print ('Draft id: %s\nDraft message: %s' % (draft['id'], draft['message']))

    return draft
  except errors.HttpError as error:
    print ('An error occurred: %s' % error)
    return None


def send_message(service, user_id, message):
  """Send an email message.

  Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

  Returns:
    Sent Message.
  """
  try:
    message = (service.users().messages().send(userId=user_id, body=message)
               .execute())
    print ('Message Id: %s' % message['id'])
    return message
  except errors.HttpError as error:
    print ('An error occurred: %s' % error)


def create_message(sender, to, subject, message_text):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}


def create_message_with_attachment(sender, to, subject, message_text, file_attachments):
  """Create a message for an email.

  Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.
    file: The path to the file to be attached.

  Returns:
    An object containing a base64url encoded email object.
  """
  message = MIMEMultipart()
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject

  msg = MIMEText(message_text)
  message.attach(msg)

  for attachment in file_attachments:

    content_type, encoding = mimetypes.guess_type(file)

    if content_type is None or encoding is not None:
      content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)
    if main_type == 'text':
      fp = open(file, 'rb')
      msg = MIMEText(fp.read(), _subtype=sub_type)
      fp.close()
    elif main_type == 'image':
      fp = open(file, 'rb')
      msg = MIMEImage(fp.read(), _subtype=sub_type)
      fp.close()
    elif main_type == 'audio':
      fp = open(file, 'rb')
      msg = MIMEAudio(fp.read(), _subtype=sub_type)
      fp.close()
    else:
      fp = open(file, 'rb')
      msg = MIMEBase(main_type, sub_type)
      msg.set_payload(fp.read())
      fp.close()
    filename = os.path.basename(file)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

  return {'raw': base64.urlsafe_b64encode(message.as_string().encode()).decode()}



def start_emailer(EMAIL_TO, SUBJECT, EMAIL_CONTENT):
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token_gmail.pickle'):
        with open('token_gmail.pickle', 'rb') as token:
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
        with open('token_gmail.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)
    
    EMAIL_CONTENT = docx2txt.process("temp.docx")


    
    EMAIL_FROM = 'me'
    EMAIL_TO = 'rupress.karen.li@gmail.com'
    EMAIL_SUBJECT = 'Hello  from Karen'
    
    message = create_message(EMAIL_FROM, EMAIL_TO, EMAIL_SUBJECT, EMAIL_CONTENT)
    create_draft(service, "me", message) ## Success!
    #print(get_messages(service, "me"))
    #read_email(service, "me", "Gamma")
    #(service, "me", message) ## Success!


    #read_email(service, "me", 'subject')


if __name__ == '__main__':
    main()
