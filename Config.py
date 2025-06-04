from datetime import datetime
import os
import shutil
import uuid
from urllib import parse as url_parse
from dotenv import load_dotenv

load_dotenv()

def _validate_env_vars():
        """
        Validates the required environment variables.
        """
        required_vars = [
            "CLIENT_ID",
            "CLIENT_SECRET",
            "REDIRECT_URI",
            "SCOPE",
            "TENANT_ID",
            "REFRESH_TOKEN",
            "ACCESS_TOKEN",
        ]

        for var in required_vars:
            if not os.getenv(var):
                raise ValueError(f"Environment variable {var} is not set.")
        return

class FileNotInRequiredFormat(Exception):
    def __init__(self, message="File is not in the required format."):
        self.message = message
        super().__init__(self.message)

class BaseConfig:
    """
    A class to load environment variables for OAuth 2.0 authentication.
    """
    def __init__(self):
        """
        Sets the environment variables for client_id, client_secret, and redirect_uri.
        """

        _validate_env_vars()
        base_url = os.getenv("BASE_URL")
        self.base_url = base_url
        self.tenant_id = os.getenv("TENANT_ID")
        self.token_url = f"{base_url}/oauth2/v2.0/token"
        self.scope = os.getenv("SCOPE")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.client_id = os.getenv("CLIENT_ID")
        self.redirect_uri = os.getenv("REDIRECT_URI") 
        self.add_token_response_to_file = os.getenv("ADD_TOKEN_RESPONSE_TO_FILE", "false").lower() == "true"
        self.auth_url = self._build_auth_url()
    
    def _build_auth_url(self):
        data = {
            'client_id': self.client_id,
            'response_type': 'code',
            'redirect_uri': self.redirect_uri,
            'response_mode': 'query',
            'scope': self.scope,
            'prompt': 'login', # force login prevent reuse of the same authcode
            "state": str(uuid.uuid4()),
            "max_age": "0"
        }

        return f"{os.getenv('BASE_URL')}/oauth2/v2.0/authorize?{url_parse.urlencode(data)}"

class TokenConfig:
    """
    A class to load environment variables for OAuth 2.0 token management.
    """
    def __init__(self):
        """
        Sets the environment variables for client_id, client_secret, and redirect_uri.
        """
        _validate_env_vars()
        self.client_id = os.getenv("CLIENT_ID")
        self.client_secret = os.getenv("CLIENT_SECRET")
        self.redirect_uri = os.getenv("REDIRECT_URI") 
        self.token_url = os.getenv("TOKEN_URL")
        self.scope = os.getenv("SCOPE")
        self.tenant_id = os.getenv("TENANT_ID")
        self.add_token_response_to_file = os.getenv("ADD_TOKEN_RESPONSE_TO_FILE", "false").lower() == "true"
        self.refresh_token = os.getenv("REFRESH_TOKEN")
        self.access_token = os.getenv("ACCESS_TOKEN")
        self.expires_in = os.getenv("EXPIRES_IN")
        self.token_type = os.getenv("TOKEN_TYPE")
        self.scope = os.getenv("SCOPE")
        self.app_name = os.getenv("APP_NAME")
        self.working_dir = os.getcwd()
        
    def __str__(self):
        return f"TokenConfig(client_id={self.client_id}, client_secret={self.client_secret}, redirect_uri={self.redirect_uri}, token_url={self.token_url}, scope={self.scope}, tenant_id={self.tenant_id}, add_token_response_to_file={self.add_token_response_to_file}, refresh_token={self.refresh_token}, access_token={self.access_token}, expires_in={self.expires_in}, token_type={self.token_type})"
    
    def test_access_token(self):
        from OutlookConnection import OutlookConnection
        OutlookConnection({
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "tenant_id": self.tenant_id,
            "add_token_to_file": self.add_token_response_to_file,
            "scope": self.scope,
            "token_url": self.token_url,
            "refresh_token": self.refresh_token,
            "access_token": self.access_token
        })._test_connection()

class OverWriteEnv:
    """
    A class to overwrite the environment variables with the token file.
    """
    def __init__(self, token_file_name:str, token_file_path:str):

        if not token_file_name.endswith('.json'):
            raise FileNotInRequiredFormat("File not in the JSON format")

        self.token_file_name = token_file_name
        self.token_file_path = os.getcwd().replace("\\", "/") if not token_file_path else token_file_path
        self.token_file_full_path = f"{token_file_path}/{token_file_name}"
        self.env_archive_path = f"{os.getcwd().replace("\\", "/")}/old_envs"

    def timestamp(self):
        return datetime.now().strftime("%Y%m%d%H%M%S")

    def load_token_details(self):
        """
        Load token details from a JSON file.
        this is done for local testing purposes only.
        In production, you should use a secure vault or environment variables.
        """
        import json
        with open(f'{self.token_file_name}', 'r') as f:
            token_data = json.load(f)
            return token_data


    def run(self):
        
        if not os.path.exists(f"{self.token_file_path}/{self.token_file_name}"):
            raise FileNotFoundError(f"File not present at the specified path. FileName: {self.token_file_name}, FilePath: {self.token_file_path}")
        
        print("file_present")
        
        if not os.path.exists(f'{self.env_archive_path}'):
            os.makedirs(f'{self.env_archive_path}', exist_ok=True)    
        
        env_path = self.env_archive_path.replace('/old_envs', '/.env')
        if os.path.exists(env_path):
            shutil.move(f"{env_path}", f"{self.env_archive_path}/env_{self.timestamp()}")

        token_details = self.load_token_details()

        with open(env_path, 'w+') as f:
            f.write(f"CLIENT_ID={os.getenv("CLIENT_ID")}\n")
            f.write(f"CLIENT_SECRET={os.getenv("CLIENT_SECRET")}\n")
            f.write(f"TENANT_ID={os.getenv("TENANT_ID")}\n")
            f.write(f"SCOPE={os.getenv('SCOPE')}\n")
            f.write(f"REFRESH_TOKEN={token_details['refresh_token']}\n")
            f.write(f"ACCESS_TOKEN={token_details['access_token']}\n")
            f.write(f"BASE_URL=https://login.microsoftonline.com/consumers\n")
            f.write(f"REDIRECT_URI={os.getenv('REDIRECT_URI')}\n")
            f.write("ADD_TOKEN_RESPONSE_TO_FILE=True\n")
