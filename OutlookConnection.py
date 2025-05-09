from datetime import datetime
import json
import os
import requests
import logging

logging.basicConfig(level=logging.INFO)

class OutlookConnection:

    def __init__(self, outlook_connection_var_name):
        self.tenant_id = outlook_connection_var_name.get("TENANT_ID")
        self.access_token = outlook_connection_var_name.get('access_token')
        self.refresh_token = outlook_connection_var_name.get('refresh_token')
        self.client_id = outlook_connection_var_name.get('client_id')
        self.client_secret = outlook_connection_var_name.get('client_secret')
        self.scope = outlook_connection_var_name.get('scope', "mail.read")
        self.add_token_to_file = outlook_connection_var_name.get('add_token_to_file', False)
        self.token_url = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
        self.working_dir = os.getcwd()

    def timestamp(self):
        return datetime.now().strftime("%Y%m%d%H%M%S")

    def _archive_old_token_details(self):
        try:
            with open('token_details.json', 'r') as f:
                old_token_data = json.load(f)
        except FileNotFoundError:
            logging.warning("No previous token details found.")
            return
        if not os.path.exists(f'{self.working_dir}\\old_tokens'):
            os.makedirs(f'{self.working_dir}\\old_tokens', exist_ok=True)    
        with open(f'{self.working_dir}\\old_tokens\\token_details_{self.timestamp()}.json', 'w') as f:
            json.dump(old_token_data, f, indent=4)
            logging.info("Old token details archived to old_token_details.json")
        return

    def _update_token_details(self, token_data):
        if not self.add_token_to_file:
            logging.info("Token details not saved to file. To save, set add_token_to_file to True.")
            return
        self.access_token = token_data.get('access_token')
        self.refresh_token = token_data.get('refresh_token')
        self.expires_in = token_data.get('expires_in')
        self.token_type = token_data.get('token_type')
        self.scope = token_data.get('scope')

        # adding timestamp and other details to token_data
        token_data['timestamp'] = self.timestamp()
        token_data['client_id'] = self.client_id
        token_data['client_secret'] = self.client_secret
        token_data['tenant_id'] = os.getenv("TENANT_ID")

        self._archive_old_token_details()
        with open('token_details.json', 'w') as f: 
            json.dump(token_data, f, indent=4)
        logging.info("Token details updated and saved to token_details.json")
        return

    def _refresh_access_token(self):
        if not all([self.refresh_token, self.client_id, self.client_secret]):
            raise Exception("Missing parameters for token refresh.")
        
        payload = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'grant_type': 'refresh_token',
            'refresh_token': self.refresh_token,
            'scope': self.scope,
        }
        print(payload) 

        response = requests.post(self.token_url, data=payload)
        if response.status_code == 200:
            token_data = response.json()
            self.access_token = token_data.get('access_token')
            self.refresh_token = token_data.get('refresh_token')
            logging.info("Access token refreshed.")
            return self._update_token_details(token_data)
        logging.error("Failed to refresh access token.")
        logging.error(response.text)
        raise Exception("Token refresh failed.")

    def _test_connection(self, max_retry_count=3, retry_count=1):

        url = "https://graph.microsoft.com/v1.0/me/messages?$top=1"
        headers = self.get_headers()
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            logging.info("Connection test successful.")
            return True
        elif response.status_code == 401:
            if retry_count > max_retry_count:
                logging.error("Max retry count reached. Unable to refresh access token.")
                raise Exception("Access token expired and can not generate new access token.")
            logging.warning(f"Access token expired. Attempting to refresh. attempt: {retry_count}")
            self._refresh_access_token()
            return self._test_connection(retry_count=retry_count + 1)
        else:
            logging.error(f"Connection failed with status code: {response.status_code}")
            logging.error(response.text)
            return False
    
    def get_connection(self, force_refresh=False):
        if force_refresh:
            # should be used when you want to force refresh the token and not use the existing one
            # common scenario is when you updated the permission on of the user
            logging.info("Force refreshing access token.")
            self._refresh_access_token()
        if not self._test_connection():
            raise Exception("Failed to establish a connection to Outlook API.")
        session = requests.Session()
        session.headers.update(self.get_headers())
        return session
    
    def get_headers(self):
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json'
        }
