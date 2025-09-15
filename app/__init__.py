import os
import csv
import sys
import time
import base64
import logging
from datetime import datetime
from io import BytesIO
from textwrap import wrap

from PIL import Image, ImageDraw, ImageFont
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders
from base64 import urlsafe_b64encode

from app.config import Config, ConfigError
from app.mail.authentication import authenticate

# Constants
RECIPIENTS_CSV = 'recipients.csv'
LAST_SENT_FILE = 'last_sent.txt'
LOG_DIR = 'log'
GIF_TEMPLATE_PATH = 'static/template.gif'
FONT_PATH = 'static/arial.ttf'

# Ensure log directory exists
if not os.path.exists(LOG_DIR):
  os.mkdir(LOG_DIR)

# Configure logging
logging.basicConfig(
  level=logging.DEBUG,
  format="%(asctime)s [%(levelname)s] %(message)s",
  handlers=[
    logging.FileHandler(f"{LOG_DIR}/LOG_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"),
    logging.StreamHandler()
  ]
)

def load_config():
  try:
    config = Config("config.json")
    logging.info("Config loaded.")
    return config
  except ConfigError as e:
    logging.error(f"Configuration Error: {e}")
    sys.exit(1)

def load_recipients(csv_file):
  if not os.path.exists(csv_file):
    logging.error(f"Recipients file '{csv_file}' not found.")
    sys.exit(1)

  recipients = []
  with open(csv_file, newline='', encoding='utf-8') as csvfile:
    reader = csv.DictReader(csvfile)
    required_headers = {"First Name", "Last Name", "Email", "Phone", "Address", "Profession", "Stage", "Industry", "LinkedIn"}
    if not required_headers.issubset(reader.fieldnames):
      logging.error(f"CSV file missing required headers. Required headers: {required_headers}")
      sys.exit(1)

    for row in reader:
      recipients.append(row)
  logging.info(f"Loaded {len(recipients)} recipients from '{csv_file}'.")
  return recipients

def get_last_sent():
  if os.path.exists(LAST_SENT_FILE):
    with open(LAST_SENT_FILE, 'r') as f:
      last_email = f.read().strip()
      if last_email:
        logging.info(f"Last sent email retrieved: {last_email}")
        return last_email
      else:
        logging.info(f"Last sent file '{LAST_SENT_FILE}' is empty.")
        return None
  return None

def set_last_sent(email):
  with open(LAST_SENT_FILE, 'w') as f:
    f.write(email)
  logging.debug(f"Set last sent email to: {email}")

def generate_funny_image(recipient):
  """
  Generates a customized GIF with embedded text for the recipient.
  Returns the binary data of the GIF.
  """
  try:
    with Image.open(GIF_TEMPLATE_PATH) as img:
      frames = []
      text = f"Look at you, {recipient['First Name']}! Working harder than Sparky to be sustainable!"
      font_size = 24  # Adjust as needed
      font = ImageFont.truetype(FONT_PATH, font_size)

      try:
        transparency = img.info.get("transparency")
        while True:
          frame = img.convert('RGBA')
          draw = ImageDraw.Draw(frame)

          # Wrap text for readability
          wrapped_text = "\n".join(wrap(text, width=30))

          # Calculate text placement
          text_width, text_height = draw.multiline_textsize(wrapped_text, font=font)
          text_x = (img.size[0] - text_width) // 2
          text_y = (img.size[1] - text_height) // 2

          draw.multiline_text((text_x, text_y), wrapped_text, fill=(255, 255, 255), font=font)

          # Convert back to palette-based mode for GIF
          frame = frame.convert('P', palette=Image.ADAPTIVE)
          frames.append(frame)

          img.seek(img.tell() + 1)
      except EOFError:
        pass  # End of frames

      # Save all frames to a buffer
      buffer = BytesIO()
      if len(frames) > 1:
        frames[0].save(
          buffer,
          format='GIF',
          save_all=True,
          append_images=frames[1:],
          loop=0,
          duration=img.info.get('duration', 100),
          transparency=transparency,
          disposal=2
        )
      else:
        frames[0].save(buffer, format='GIF', transparency=transparency)

      buffer.seek(0)
      gif_data = buffer.read()
      logging.debug(f"Generated GIF for {recipient['Email']}.")
      return gif_data
  except Exception as e:
    logging.error(f"Failed to generate funny image for {recipient['Email']}: {e}")
    raise

def generate_email_body(recipient, image_cid):
  """
  Generates the HTML body of the email, embedding the GIF via CID.
  """
  try:
    with open('template.html', 'r', encoding='utf-8') as f:
      template = f.read()
    
    dynamic_image = f'<img src="cid:{image_cid}" alt="Sparky doing push-ups">'
    
    body = template.format(
      first_name=recipient['First Name'],
      email=recipient['Email'],
      dynamic_image=dynamic_image
    )
    return body
  except Exception as e:
    logging.error(f"Failed to load email template: {e}")
    raise

def build_message(destination, subject, body, gif_data, gif_cid, attachments=None, config=None):
  """
  Builds a MIME message with embedded GIF and optional attachments.
  """
  if config is None:
    raise ValueError("A config instance must be provided to build_message.")

  # Create the root message with 'related' subtype to handle embedded images
  msg_root = MIMEMultipart('related')
  msg_root['To'] = destination
  msg_root['From'] = config['sender_email']
  msg_root['Subject'] = subject

  # Create the alternative part for HTML
  msg_alternative = MIMEMultipart('alternative')
  msg_root.attach(msg_alternative)

  # Attach the HTML body
  msg_text = MIMEText(body, 'html')
  msg_alternative.attach(msg_text)

  # Attach the GIF image
  if gif_data and gif_cid:
    msg_image = MIMEImage(gif_data, _subtype="gif")
    msg_image.add_header('Content-ID', f'<{gif_cid}>')
    msg_image.add_header('Content-Disposition', 'inline', filename='funny_sparky.gif')
    msg_root.attach(msg_image)

  # Attach any additional files
  if attachments:
    for filename in attachments:
      add_attachment(msg_root, filename)

  # Encode the message
  raw_message = urlsafe_b64encode(msg_root.as_bytes()).decode()
  return {'raw': raw_message}

def add_attachment(message, filename):
  """
  Attaches a file to the email message.
  """
  try:
    path = os.path.join('attachments', filename)
    with open(path, 'rb') as f:
      part = MIMEBase('application', 'octet-stream')
      part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
    message.attach(part)
    logging.debug(f"Attached file {filename}.")
  except Exception as e:
    logging.error(f"Failed to attach file {filename}: {e}")
    raise

def send_message(destination, subject, body, gif_data, gif_cid, attachments=None, config=None):
  """
  Sends the email message via the authenticated mail service.
  """
  if config is None:
    raise ValueError("A config instance must be passed to send_message.")

  logging.info(f"Sending message to {destination}")
  try:
    mail_service = authenticate()
    message_body = build_message(
      destination=destination,
      subject=subject,
      body=body,
      gif_data=gif_data,
      gif_cid=gif_cid,
      attachments=attachments,
      config=config
    )
    sent_message = mail_service.users().messages().send(userId="me", body=message_body).execute()
    logging.info(f"Message sent to {destination} with Message ID: {sent_message.get('id')}")
    return sent_message
  except Exception as e:
    logging.error(f"Failed to send message to {destination}: {e}")
    raise

def start():
  logging.info("Starting the email sending process...")
  config = load_config()
  recipients = load_recipients(RECIPIENTS_CSV)

  try:
    subject = config["subject"]
  except KeyError:
    logging.error("Config missing 'subject' key.")
    sys.exit(1)

  try:
    test_mode = config["test"]
  except KeyError:
    logging.error("Config missing 'test' key.")
    sys.exit(1)

  try:
    test_email = config["test_email_recipient"]
  except KeyError:
    logging.error("Config missing 'test_email_recipient' key.")
    sys.exit(1)

  last_sent_email = get_last_sent()
  start_index = 0

  if last_sent_email:
    for i, recipient in enumerate(recipients):
      if recipient['Email'] == last_sent_email:
        start_index = i + 1
        break

  total_recipients = len(recipients)
  if start_index >= total_recipients:
    logging.info("All emails have been sent already.")
    return

  logging.info(f"Preparing to send emails to {total_recipients - start_index} recipients.")

  if test_mode:
    if not test_email:
      logging.error("Test mode is enabled but no test_email_recipient is specified in config.")
      sys.exit(1)
    test_recipient = recipients[0]
    try:
      # Generate GIF data
      gif_data = generate_funny_image(test_recipient)
      # Define a unique Content-ID for the GIF
      gif_cid = "funny_image"
      # Generate email body with CID reference
      body = generate_email_body(test_recipient, gif_cid)
      # Send the test email
      send_message(
        destination=test_email,
        subject=subject,
        body=body,
        gif_data=gif_data,
        gif_cid=gif_cid,
        attachments=None,
        config=config
      )
      logging.info(f"Test email sent to {test_email}.")
    except Exception as e:
      logging.error(f"Failed to send test email to {test_email}: {e}")
    return

  for i in range(start_index, total_recipients):
    recipient = recipients[i]
    try:
      # Generate GIF data
      gif_data = generate_funny_image(recipient)
      # Define a unique Content-ID for the GIF
      gif_cid = "funny_image"
      # Generate email body with CID reference
      body = generate_email_body(recipient, gif_cid)
      # Send the email
      send_message(
        destination=recipient['Email'],
        subject=subject,
        body=body,
        gif_data=gif_data,
        gif_cid=gif_cid,
        attachments=None,
        config=config
      )
      logging.info(f"Email {i + 1}/{total_recipients} sent to {recipient['Email']}.")
      set_last_sent(recipient['Email'])

      # Delay to adhere to the sending limits
      if (i - start_index + 1) % 6 == 0:  # After every 6 emails
        logging.debug("Hourly limit reached. Waiting for 1 hour...")
        time.sleep(3600)  # Wait for 1 hour
      else:
        logging.debug("Waiting 10 minutes (600 seconds) before sending the next email...")
        time.sleep(600)  # Wait for 600 seconds between emails

    except Exception as e:
      logging.error(f"Failed to send email to {recipient['Email']}: {e}")
      break

  logging.info("Email sending process completed.")

if __name__ == "__main__":
  start()
