import os
import sys
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/script.projects", "https://www.googleapis.com/auth/spreadsheets"]

FUNCTION_NAME = "createFuelManagementTemplate"

def main():
  """Executes an Apps Script function."""
  if len(sys.argv) < 2:
    print("Usage: python3 run_script.py <SCRIPT_ID>")
    sys.exit(1)

  script_id = sys.argv[1]
  creds = None

  if not os.path.exists("creds.json"):
      print("Error: creds.json not found.")
      print("Please follow the instructions to create your OAuth 2.0 credentials and save the file as creds.json in the same directory as this script.")
      sys.exit(1)


  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file("creds.json", SCOPES)
      creds = flow.run_local_server(port=0)

    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("script", "v1", credentials=creds)

    # Create an execution request object.
    request = {"function": FUNCTION_NAME}

    # Make the API request.
    print(f"Executing function '{FUNCTION_NAME}' in script {script_id}...")
    response = service.scripts().run(scriptId=script_id, body=request).execute()

    if "error" in response:
      # The API executed, but the script returned an error.
      error = response["error"]["details"][0]
      print(f"Script error message: {error['errorMessage']}")
      if "scriptStackTraceElements" in error:
        print("Script error stacktrace:")
        for trace in error["scriptStackTraceElements"]:
          print(f"\t{trace['function']}: {trace['lineNumber']}")
    else:
      print("Script executed successfully!")
      if "response" in response and "result" in response["response"]:
          print("Script response:")
          print(response["response"]["result"])

  except HttpError as error:
    print(f"An error occurred: {error}")
    print(error.content)
  except Exception as e:
    print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
  main()