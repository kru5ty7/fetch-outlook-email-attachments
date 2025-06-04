from OutlookConnection import OutlookConnection
import logging

from dotenv import load_dotenv
load_dotenv()

logging.basicConfig(level=logging.INFO)

class OutlookEmail:

    def __init__(self, outlook_connection_var_name, force_refresh=False):
        self.outlook_connection = OutlookConnection(outlook_connection_var_name)
        self.session = self.outlook_connection.get_connection(force_refresh)
        self.category = self.get_categoryId_by_name('marked read by integration')
        self.category_name = "marked read by integration"

    def get_unread_emails_from_folder(self, folder="inbox"):
        """
        Fetch unread emails from the folder default `inbox`.
        Returns a list of unread emails.
        """

        folder_id = self.get_folderId_by_name(folder)
        url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages?$filter=isRead eq false"
        response = self.session.get(url)

        if response.status_code == 200:
            messages = response.json().get('value', [])
            return messages
        else:
            logging.error(f"Failed to fetch unread emails. Status code: {response.status_code}, Response: {response.text}")
            print(response.text)
            return []

    def get_email_details(self, email_id):
        """
        Fetch details of a specific email by its ID.
        Returns the email details.
        """
        session = self.session
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"
        response = session.get(url)

        if response.status_code == 200:
            email_details = response.json()
            return email_details
        else:
            logging.error(f"Failed to fetch email details. Status code: {response.status_code}, Response: {response.text}")
            return None

    def mark_email_as_read(self, email_id):
        """
        Mark an email as read by its ID.
        Returns True if successful, False otherwise.
        """
        session = self.session
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"
        headers = {
            'Content-Type': 'application/json'
        }
        data = {
            "isRead": True
        }
        response = session.patch(url, headers=headers, json=data)
        # self.update_email_category(email_id)
        if response.status_code == 200:
            return True
        else:
            logging.error(f"Failed to mark email as read. Status code: {response.status_code}, Response: {response.text}")
            return False
    
    def mark_email_as_unread(self, email_id):
        """
        Mark an email as unread by its ID.
        Returns True if successful, False otherwise.
        """
        session = self.session
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"
        headers = {
            'Content-Type': 'application/json'
        }
        data = {
            "isRead": False
        }
        response = session.patch(url, headers=headers, json=data)

        if response.status_code == 200:
            return True
        else:
            logging.error(f"Failed to mark email as unread. Status code: {response.status_code}, Response: {response.text}")
            return False
        
    def mark_email_as_unread_batch(self, email_ids):
        """
        Mark multiple emails as unread by their IDs.
        Returns True if successful, False otherwise.
        """
        session = self.session
        headers = {
            'Content-Type': 'application/json'
        }
        failed_emails = []
        for email_id in email_ids:
            url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"
            data = {
                "isRead": False
            }
            response = session.patch(url, headers=headers, json=data)

            if response.status_code != 200:
                logging.error(f"Failed to mark email {email_id} as unread. Status code: {response.status_code}, Response: {response.text}")
                failed_emails.append(email_id)
        if failed_emails:
            logging.error(f"Failed to mark the following emails as unread: {failed_emails}")
            return False
        return True

    def get_email_attachments(self, email_id):
        """
        Fetch attachments of a specific email by its ID.
        Returns a list of attachments.
        """
        session = self.session
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/attachments"
        response = session.get(url)

        if response.status_code == 200:
            attachments = response.json().get('value', [])
            return attachments
        else:
            logging.error(f"Failed to fetch email attachments. Status code: {response.status_code}, Response: {response.text}")
            return []
        
    def _get_all_categories(self):
        """
        Fetch all categories.
        Returns a list of categories.
        """
        url = "https://graph.microsoft.com/v1.0/me/outlook/masterCategories"
        response = self.session.get(url)

        if response.status_code == 200:
            categories = response.json().get('value', [])
            return categories
        else:
            logging.error(f"Failed to fetch categories. Status code: {response.status_code}, Response: {response.text}")
            return []

    def get_categoryId_by_name(self, category_name:str):
        """
        Get the category ID by its name.
        Returns the category ID if found, None otherwise.
        """
        
        categories = self._get_all_categories()

        for category in categories:
            if category['displayName'].lower() == category_name.lower():
                return category['id']
        logging.error(f"Category '{category_name}' not found.")
        return None

    def _get_all_folders(self):
        """
        Fetch all mail folders.
        Returns a list of folders.
        """
        url = "https://graph.microsoft.com/v1.0/me/mailFolders"
        response = self.session.get(url)

        if response.status_code == 200:
            folders = response.json().get('value', [])
            self.all_email_folders = folders
            logging.info(f"Fetched all mail folders.")
            return folders
        else:
            logging.error(f"Failed to fetch mail folders. Status code: {response.status_code}, Response: {response.text}")
            return []


    def get_folderId_by_name(self, folder_name:str):
        """
        Get the folder ID by its name.
        Returns the folder ID if found, None otherwise.
        """
        
        self._get_all_folders()

        for folder in self.all_email_folders:
            if folder['displayName'].lower() == folder_name.lower():
                return folder['id']
        logging.error(f"Folder '{folder_name}' not found.")
        return None

    def get_emails_from_folder(self, folder="inbox"):
        """
        Fetch emails from a specific folder.
        Returns a list of emails.
        """
        folder_id = self.get_folderId_by_name(folder)

        url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages"
        response = self.session.get(url)

        if response.status_code == 200:
            messages = response.json().get('value', [])
            return messages
        else:
            logging.error(f"Failed to fetch emails from {folder}. Status code: {response.status_code}, Response: {response.text}")
            return []
 
    def get_attachment_content(self, email_id, attachment_id):
        """
        Fetch the content of a specific attachment by its ID.
        Returns the attachment content.
        """
        session = self.session
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/attachments/{attachment_id}/$value"
        response = session.get(url)

        if response.status_code == 200:
            return response.content
        else:
            logging.error(f"Failed to fetch attachment content. Status code: {response.status_code}, Response: {response.text}")
            return None

    def update_email_category(self, email_id):
        # Does not work yet
        """
        Update the category of an email by its ID.
        Returns True if successful, False otherwise.
        """
        session = self.session
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"
        headers = {
            'Content-Type': 'application/json'
        }
        data = {
            "singleValueExtendedProperties": [
                {
                    "id": f"String {'{'}{self.category}{'}'} Name Color",
                    "value": "Green"
                }
            ]
        }
        response = session.patch(url, headers=headers, json=data)

        if response.status_code == 200:
            logging.info(f"Email {email_id} updated with category '{self.category_name}'.")
            return True
        else:
            logging.error(f"Failed to update email category. Status code: {response.status_code}, Response: {response.text}")
            return False
    
    def search_email_in_draft(self, search_param, search_param_value):

        if not isinstance(search_param, str):
            logging.error(f"Excepted search params to be of type `str`. Received of type {type(search_param)}")
            raise TypeError(f"Excepted search params to be of type `str`. Received of type {type(search_param)}")

        if search_param not in ['id', 'subject']:
            logging.error(f"Excepted search params `id` or `subject`. Received {search_param}")
            raise ValueError(f"Excepted search params `id` or `subject`. Received {search_param}")

        draft_emails = self.get_emails_from_folder(folder="drafts")

        for draft in draft_emails:
            if draft[search_param] == search_param_value:
                logging.info(f"Email found in draft status with search_param {search_param} and value {search_param_value}")
                return draft

        logging.exception(f"No email was found in draft status with search_param {search_param} and value {search_param_value}")
        return {}


    def draft_email(self, subject, body, toRecipients, ccRecipients=[], bccRecipients=[], email_id = None, update=False):
        """
        Draft and email
        """
        session = self.session
        if update == True:
            url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"
        else:
            url = "https://graph.microsoft.com/v1.0/me/messages"
        data = {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": body
            },
            "toRecipients": [ 
                {
                    "emailAddress": {
                        "address": toRecipient,
                    }
                }    
            for toRecipient in toRecipients ],
            "ccRecipients": [ 
                {
                    "emailAddress": {
                        "address": ccRecipient,
                    }
                }    
            for ccRecipient in ccRecipients ],
            "bccRecipients": [ 
                {
                    "emailAddress": {
                        "address": bccRecipient,
                    }
                }    
            for bccRecipient in bccRecipients ]
        }
        
        if update == True:
            response = session.patch(url, json=data)
        else:
            response = session.post(url, json=data)
        resp_json = response.json()
        if response.status_code == 201 or response.status_code == 200:
            logging.info(f"Successfully drafted / updated email with subject {subject}. Email ID: {resp_json['id']}")
            return {
                "status": "success",
                "draft_details": resp_json
            }
        else:
            logging.error(f"Failed to create / update email Draft. Status code: {response.status_code}, Response: {response.text}")
            return {
                "status": "fail",
                "draft_details": resp_json
            }
        
    def update_draft(self, create_draft_if_not_found, subject, body, toRecipients, email_id=None, ccRecipients=[], bccRecipients=[]):
        draft_email = None
        if email_id:
            logging.info(f"Searching Drafts for email with ID {email_id}")
            draft_email = self.search_email_in_draft('id', email_id)
        if not draft_email and subject:
            logging.info(f"Searching Drafts for email with subject {subject}")
            draft_email = self.search_email_in_draft('subject', subject)

        if not draft_email:
            if create_draft_if_not_found:
                logging.info(f"No email found for the ID / Subject. Initiated new draft creation. Set `create_draft_if_not_found` to avoid new draft creation")
                if subject and body and toRecipients:
                    draft_email = self.draft_email(subject, body, toRecipients, ccRecipients, bccRecipients)
                    return {
                        "draft_updated": False,
                        "draft_created": True,
                        "response": draft_email,
                        "message": "No draft found with provided details and a new draft created with details"
                    }
                return {
                        "draft_updated": False,
                        "draft_created": False,
                        "response": {},
                        "message": "No draft found with provided details and all mandatory fields are not present for a new draft creation"
                    }
            return {
                        "draft_updated": False,
                        "draft_created": False,
                        "response": {},
                        "message": "No draft found with provided details and new draft creation skipped"
                    }
        
        return self.draft_email(subject, body, toRecipients, ccRecipients, bccRecipients, draft_email['id'], True)

    def send_email(self, email_id=None, subject=None):
        draft_email = None
        if email_id:
            logging.info(f"Searching Drafts for email with ID {email_id}")
            draft_email = self.search_email_in_draft('id', email_id)
        if not draft_email and subject:
            logging.info(f"Searching Drafts for email with subject {subject}")
            draft_email = self.search_email_in_draft('subject', subject)
        if not draft_email:
            return {
                "sent_email": False,
                "status": 404,
                "response": {},
                "message": "No draft email found with provided details"
            }
        
        url = f"https://graph.microsoft.com/v1.0/me/messages/{draft_email['id']}/send"

        response = self.session.post(url)

        if response.status_code == 202:
            logging.info(f"Email Message sent successfully. Email Subject: {draft_email['subject']} Email ID: {draft_email['id']}")
            return {
                "sent_email": True,
                "status": 202,
                "response": response.json(),
                "message": "Successfully sent email"
            }
        else:
            logging.info(f"Failure occurred while sending email. Email Subject: {draft_email['subject']} Email ID: {draft_email['id']}. Response Code: {response.status_code}, Response Details: {response.text}")
            return {
                "sent_email": False,
                "status": response.status_code,
                "response": response.json(),
                "message": response.text
            }

    def delete_email(self, email_id):
        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}"

        response = self.session.delete(url)

        if response.status_code == 201 or response.status_code == 200:
            logging.info(f"Successfully deleted email. Email ID: {email_id}")
            return {
                "status": "success",
                "details": "Email deleted successfully"
            }
        else:
            logging.error(f"Failed to delete email. Status code: {response.status_code}, Response: {response.text}")
            return {
                "status": "fail",
                "details": response.text
            }


    def move_email_to_another_folder(self, email_id, folder_id):

        url = f"https://graph.microsoft.com/v1.0/me/messages/{email_id}/move"

        data = {
            "destinationId": folder_id
        }

        response = self.session.post(url, json=data)

        resp_json = response.json()
        if response.status_code == 201 or response.status_code == 200:
            logging.info(f"Successfully moved email. Email ID: {resp_json['id']}")
            return {
                "status": "success",
                "details": resp_json
            }
        else:
            logging.error(f"Failed to move email to different folder. Status code: {response.status_code}, Response: {response.text}")
            return {
                "status": "fail",
                "details": resp_json
            }