"""
Remote Executor for Google Apps Script

This script provides a command-line interface to execute a Google Apps Script
function remotely. It handles user authentication via the OAuth 2.0 protocol,
manages credentials, and calls the specified Apps Script function using the
Google Apps Script API.

This is designed to be used in automated workflows where manual execution of
the script from within Google Sheets is not feasible.
"""

import argparse
import os
import os.path
import sys

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Define the scope required to access Google Sheets and execute Apps Script.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


def get_credentials():
    """
    Manages OAuth 2.0 authentication for the script.

    Handles the credential flow by looking for 'token.json', refreshing expired
    tokens, or initiating a new authorization flow via 'creds.json'.

    Returns:
        google.oauth2.credentials.Credentials: The authenticated credentials object,
        or None if authentication fails.
    """
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired credentials...")
            creds.refresh(Request())
        else:
            if not os.path.exists("creds.json"):
                print(
                    "Error: 'creds.json' not found. Please follow the README instructions.",
                    file=sys.stderr,
                )
                return None

            print("Initiating new authorization flow...")
            flow = InstalledAppFlow.from_client_secrets_file("creds.json", SCOPES)
            creds = flow.run_local_server(port=0)

        with open("token.json", "w") as token:
            token.write(creds.to_json())
            print("Credentials saved to 'token.json'.")

    return creds


def main():
    """
    Executes the 'createFuelManagementTemplate' Google Apps Script function.

    Authenticates the user, builds the API service, and calls the target
    function using a Script ID provided as a command-line argument.
    """
    parser = argparse.ArgumentParser(
        description="Execute a Google Apps Script function remotely.",
        usage="python execute_script.py <SCRIPT_ID>",
        epilog="Ensure 'creds.json' is present for the initial run. See README.md for setup details."
    )
    parser.add_argument(
        "script_id",
        help="The ID of the Google Apps Script project to execute.",
    )

    if len(sys.argv) == 1:
        parser.print_help(sys.stderr)
        sys.exit(1)

    args = parser.parse_args()
    script_id = args.script_id

    creds = get_credentials()
    if not creds:
        print("Authentication failed. Exiting.", file=sys.stderr)
        return

    try:
        service = build("script", "v1", credentials=creds)

        request = {
            "function": "createFuelManagementTemplate",
            "parameters": [],
            "devMode": True,
        }

        print(f"Executing Apps Script function for script ID: {script_id}...")
        response = service.scripts().run(body=request, scriptId=script_id).execute()

        if "error" in response:
            error = response["error"]["details"][0]
            print(f"Script error message: {error['errorMessage']}", file=sys.stderr)
            if "scriptStackTraceElements" in error:
                print("Script stack trace:", file=sys.stderr)
                for trace in error["scriptStackTraceElements"]:
                    print(
                        f"\t{trace['function']}: {trace['lineNumber']}",
                        file=sys.stderr,
                    )
        else:
            print("Script executed successfully!")

    except HttpError as err:
        print(f"An API error occurred: {err}", file=sys.stderr)
    except Exception as e:
        print(f"An unexpected error occurred: {e}", file=sys.stderr)


if __name__ == "__main__":
    main()