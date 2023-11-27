from dotenv import load_dotenv
import json
import os
import requests

basedir = os.path.abspath(os.path.dirname(__file__))
load_dotenv(os.path.join(basedir, '.env'))

# Retrieve the credentials from environment variables
username = os.getenv("USERNAME")
password = os.getenv("PASSWORD")

def login():

    print(f'Trying to login in wnp 🙂......')
    login_url = f"{os.getenv('LOGIN_URL')}/signin"

    login_data = {
    "email_address": username,
    "password": password
}

    param_without_session = _createAdditionalRequestParams()
    print(f'Login In Process . . . . .')

    try:
        # Send a POST request to the login URL with the credentials
        response = requests.post(login_url, data=json.dumps(login_data), **param_without_session)
        
        # Check the response status code
        if response.status_code == 200:
            response_json = response.json()
            sessionId = response_json.get('session_id')
            # extract session id from response and send it as param to _createAdditionalRequestParams
            params_with_sessionId = _createAdditionalRequestParams(sessionId) # get the sessionId after login
            
            print(f'Login successful for wnp 😊')
            return params_with_sessionId
    except Exception as e:
         print(f'Logging error while logging in😑:{e}')
    return False


def baseUrl():
     url = os.getenv("URL_BASE")
     return url


def signout(sessionId):
    sign_out_url = f"{os.getenv('LOGIN_URL')}/user/signout"
    params = _createAdditionalRequestParams(sessionId)
    
    try:
        sign_out = requests.post(sign_out_url, **params)
        sign_out_json = sign_out.json()
        print(f"Sign out successful....🙂:{sign_out_json.get('status')}")
    except Exception as e:
        print(f'Unable to sign out due to:{e}')

def _createAdditionalRequestParams(sessionId=None):
        params = {
            'headers': {
                'ICClientKey': os.getenv("ICClientKey"),
                'content-type': 'application/json' 
                }
            }
        if sessionId:
            params['headers']['ICSessionId'] = sessionId
        return params