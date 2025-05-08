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