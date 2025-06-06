from http.server import BaseHTTPRequestHandler, HTTPServer
import requests
from urllib.parse import parse_qs, urlparse
import json
import logging
import os

from Config import BaseConfig

from dotenv import load_dotenv
load_dotenv()

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


class OAuthHandler(BaseHTTPRequestHandler):

    config = BaseConfig()
   
    def _get_code_state_from_query_params(self, query_params):
        """
        Extracts the 'code' and 'state' parameters from the query parameters.
        """
        code = query_params.get('code', [None])[0]
        state = query_params.get('state', [None])[0]

        logging.info(f"Received code: {code}")
        logging.info(f"Received state: {state}")

        return code, state
    
    def _send_response(self, code, state, response_code, response_text, retry):
        """
        Sends a response back to the client with the received code and state.
        """
        self.send_response(response_code)
        self.send_header("Content-type", "text/html")
        self.end_headers()

        # Send a simple HTML response
        response_html = f"""
            <html>
                <head><title>OAuth 2.0 Callback </title></head>
                <body>
                    <h1>OAuth 2.0 Callback</h1>
                    <p>Retry Attempt {retry+1}</p>
                    <p>Authorization code received: {code}</p>
                    <p>State received: {state}</p>
                    <p>Response code: {response_code}</p>
                    <p>Response text: {response_text}</p>
                </body>
            </html>
        """
        
        self.wfile.write(response_html.encode('utf-8'))

    def _write_token_to_file(self, token_response):
        """
        Writes the token response to a file.
        """
        # Assuming you want to write the token response to a file
        if self.config.add_token_response_to_file:
            with open("token_response.json", "w") as token_file:
                json.dump(token_response, token_file, indent=4)
                logging.info("Token response written to file.")

    def _get_token_for_auth_code(self, code:str, state:str, retry:int=0):
        """
        Handles the token request using the authorization code received from the OAuth 2.0 server.
        This function is called after the authorization code is received in the callback URL.
        """

        token_payload = {
            "client_id": self.config.client_id,
            "client_secret": self.config.client_secret,
            'scope': self.config.scope,
            "grant_type": "authorization_code",
            "code": code.strip(),
            "redirect_uri": self.config.redirect_uri,
        }

        # This is for demonstration purposes in local. In a real application, do not log sensitive information.
        logging.info(f"Sending request to {self.config.token_url} with payload {token_payload}") 

        response = requests.post(self.config.token_url, data=token_payload)

        if response.status_code == 200:
            logging.info("Token request successful.")
            token_response = response.json()
            logging.info(f"Token response: {token_response}")

            # Set the environment variables for the tokens
            self._write_token_to_file(token_response)
            os.environ["BEARER_TOKEN"] = token_response.get("access_token", None)
            try:
                os.environ["REFRESH_TOKEN"] = token_response.get("refresh_token", None)
            except KeyError:
                logging.warning("No refresh token found in the response.")
                raise KeyError("No refresh token found in the response.")

            os.environ["TOKEN_TYPE"] = token_response.get("token_type", None)

            self._send_response(code, state, 200, "Token request successful.", retry)
            return token_response   
        else:
            logging.error(f"Token request failed with status code {response.status_code}.")
            logging.error(f"Response: {response.text}")
            if retry < 3:

                logging.error(f"retrying..{retry}")
                self._get_token_for_auth_code(code, state, retry+1)
            
            self._send_response(code, state, response.status_code, response.text, retry)

    def do_GET(self):
        # Parse the query parameters from the URL
        parsed_path = urlparse(self.path)
        query_params = parse_qs(parsed_path.query)

        code, state = self._get_code_state_from_query_params(query_params)

        if code is None or state is None:
            logging.error("Missing 'code' or 'state' in the query parameters.")
            self._send_response(None, None, 400, "Missing 'code' or 'state' in the query parameters.", 0)
            exit(1)
            return

        if code:
            self._get_token_for_auth_code(code, state)
            exit(0)

        self._send_response(code, state, 500, "Internal Server Error", 0)
        exit(1)

class OAuthRunner:
    def __init__(self):
        self.config = BaseConfig()

    def _open_browser(self):
        """
        Opens the default web browser to the specified URL.
        """
        import webbrowser
        webbrowser.open(self.config.auth_url)

    def _run_local_server_for_oauth(self):
        httpd = HTTPServer(('localhost', 8000), OAuthHandler)
        print("Starting server on http://localhost:8000")
        httpd.serve_forever()

    def run(self):
        self._open_browser()
        self._run_local_server_for_oauth()

if __name__ == "__main__":
    oauth_runner = OAuthRunner()
    oauth_runner.run()
