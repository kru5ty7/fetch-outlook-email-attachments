import os
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
        self.working_dir = os.getcwd()
        
    def __str__(self):
        return f"TokenConfig(client_id={self.client_id}, client_secret={self.client_secret}, redirect_uri={self.redirect_uri}, token_url={self.token_url}, scope={self.scope}, tenant_id={self.tenant_id}, add_token_response_to_file={self.add_token_response_to_file}, refresh_token={self.refresh_token}, access_token={self.access_token}, expires_in={self.expires_in}, token_type={self.token_type})"