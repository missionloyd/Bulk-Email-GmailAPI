from base64 import urlsafe_b64encode
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from mimetypes import guess_type as guess_mime_type
import os

def add_attachment(message, filename):
    content_type, encoding = guess_mime_type(filename)
    if content_type is None or encoding is not None:
        content_type = 'application/octet-stream'
    main_type, sub_type = content_type.split('/', 1)

    with open(filename, 'rb') as fp:
        file_data = fp.read()

    if main_type == 'text':
        msg = MIMEText(file_data.decode(), _subtype=sub_type)
    else:
        from email.mime.base import MIMEBase
        msg = MIMEBase(main_type, sub_type)
        msg.set_payload(file_data)

    filename = os.path.basename(filename)
    msg.add_header('Content-Disposition', 'attachment', filename=filename)
    message.attach(msg)

def build_message(destination, subject, body, attachments=None, config=None):
    if config is None:
        raise ValueError("A config instance must be provided to build_message.")

    if not attachments:
        message = MIMEText(body, 'html')
    else:
        message = MIMEMultipart()
        message.attach(MIMEText(body, 'html'))
        for filename in attachments:
            add_attachment(message, filename)

    message['to'] = destination
    message['from'] = config['sender_email']
    message['subject'] = subject

    return {'raw': urlsafe_b64encode(message.as_bytes()).decode()}
