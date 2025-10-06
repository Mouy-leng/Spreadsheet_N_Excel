import os
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Define the scope required to access the Google Sheets and Apps Script APIs.
# This scope allows the script to view and manage spreadsheets that the user has
# permitted access to.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID of the Google Apps Script project to be executed.
# This ID is retrieved from an environment variable for security and flexibility.
SCRIPT_ID = os.environ.get("Script_ID")


def main():
    """
    Authenticates the user, executes a specified Google Apps Script function,
    and handles any errors that may occur.

    This function manages the OAuth 2.0 authentication flow by checking for
    existing, valid credentials in 'token.json'. If credentials are not
    available or are expired, it initiates a new authorization flow using
    'creds.json'.

    Once authenticated, it calls the 'createFuelManagementTemplate' function
    in the specified Apps Script project.
    """
    creds = None
    # The 'token.json' file stores the user's access and refresh tokens.
    # It is created automatically upon successful completion of the
    # authorization flow.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    # If there are no valid credentials, initiate the OAuth 2.0 flow.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # 'creds.json' is required to start the authorization process.
            # This file can be downloaded from the Google Cloud Console.
            if not os.path.exists("creds.json"):
                print(
                    "Error: 'creds.json' not found. Please follow the instructions in the README to obtain this file."
                )
                return

            flow = InstalledAppFlow.from_client_secrets_file("creds.json", SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the new credentials for future runs.
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        # Build the service object for the Google Apps Script API.
        service = build("script", "v1", credentials=creds)

        # Prepare the request to execute the 'createFuelManagementTemplate' function.
        request = {
            "function": "createFuelManagementTemplate",
            "parameters": [],
            "devMode": True,
        }

        # Execute the request.
        response = (
            service.scripts().run(body=request, scriptId=SCRIPT_ID).execute()
        )

        # Handle the response from the Apps Script execution.
        if "error" in response:
            error = response["error"]["details"][0]
            print(f"Script error message: {error['errorMessage']}")
            if "scriptStackTraceElements" in error:
                print("Script stack trace:")
                for trace in error["scriptStackTraceElements"]:
                    print(
                        f"\t{trace['function']}: {trace['lineNumber']}"
                    )
        else:
            print("Script executed successfully!")

    except HttpError as err:
        # Handle errors from the Google API client.
        print(f"An API error occurred: {err}")
    except Exception as e:
        # Handle other potential exceptions.
        print(f"An unexpected error occurred: {e}")


if __name__ == "__main__":
    main()