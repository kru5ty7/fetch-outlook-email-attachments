from Config import TokenConfig

from OutlookEmail import OutlookEmail
import logging

logging.basicConfig(level=logging.INFO)

config = TokenConfig()

# def load_token_details():
#     """
#     Load token details from a JSON file.
#     this is done for local testing purposes only.
#     In production, you should use a secure vault or environment variables.
#     """
#     import json
#     try:
#         with open('token_details.json', 'r') as f:
#             token_data = json.load(f)
#             return token_data
#     except FileNotFoundError:
#         logging.error("Token details file not found.")
#         return None

basic_info = {
        "client_id": config.client_id,
        "client_secret": config.client_secret,
        "tenant_id": config.tenant_id,
        "add_token_to_file": config.add_token_response_to_file,
        "scope": config.scope,
        "token_url": config.token_url,
        "refresh_token": config.refresh_token,
        "access_token": config.access_token
    }

sensor = OutlookEmail(basic_info, True)
all_emails = sensor.get_unread_emails_from_folder('xxx-xxx-xxx@outlook.com')

if not all_emails:
    logging.info("No unread emails found.")

else:
    for email in all_emails:
        if email['hasAttachments']:
            attachments = sensor.get_email_attachments(email['id'])
            for attachment in attachments:
                try:
                    attachment_name:str = attachment['name']
                    if not attachment_name.endswith('.csv'):
                        logging.info(f"Attachment {attachment_name} is not a CSV file. Skipping.")
                        sensor.mark_email_as_read(email['id'])
                        continue
                    attachment_response = sensor.get_attachment_content(email['id'], attachment['id'])
                    if not attachment_response:
                        sensor.mark_email_as_unread(email['id'])
                        continue
                    logging.info(f"Attachment Name: {attachment_name}")
                    with open(attachment_name, 'wb') as f:
                        f.write(attachment_response)
                    logging.info(f"Attachment {attachment_name} saved successfully.")
                    if not sensor.mark_email_as_read(email['id']):
                        raise Exception("Failed to mark email as read after processing attachment.")
                    logging.info(f"Email {email['id']} marked as read.")
                except Exception as e:
                    logging.error(f"Error processing attachment: {e}")
                    sensor.mark_email_as_unread(email['id'])
                    raise e




