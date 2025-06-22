
import os
from google.auth.transport import requests
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow  # type: ignore
from googleapiclient.discovery import build  # type: ignore
from googleapiclient.errors import HttpError  # type: ignore
from dotenv import load_dotenv

load_dotenv()

env_scopes = os.environ.get("SCOPES")
if env_scopes is not None:
    SCOPES = env_scopes.split(',')
else:
    raise EnvironmentError("SCOPES variable is not set")

DOCUMENT_ID = os.environ.get("DOCUMENT_ID")
SHEET_ID = os.environ.get("SHEET_ID")
SECRET = os.environ.get("SECRET")
TOKEN_JSON = os.environ.get("TOKEN_JSON")


class Creds():
    def __init__(self, token_path=TOKEN_JSON, scopes=SCOPES, secret=SECRET):
        self.token_path = token_path
        self.scopes = scopes
        self.secret = secret
        self.credentials = None

    def generated_creds(self):
        if os.path.exists(self.token_path):
            self.credentials = Credentials.from_authorized_user_file(
                self.token_path, self.scopes
            )

        if not self.credentials or not self.credentials.valid:
            if self.credentials and self.credentials.expired and self.credentials.refresh_token:
                self.credentials.refresh(requests.Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.secret, self.scopes)
                self.credentials = flow.run_local_server(port=0)

            with open(self.token_path, "w") as token:
                token.write(self.credentials.to_json())


class Sheets(Creds):
    def __init__(self, cell_range, input="", sheet_id=SHEET_ID, token_path=TOKEN_JSON, scopes=SCOPES, secret=SECRET):
        super().__init__(token_path, scopes, secret)
        self.cell_range = cell_range
        self.input = input
        self.sheet_id = sheet_id

    def read_sheet(self):
        self.generated_creds()
        try:
            service = build("sheets", "v4", credentials=self.credentials)
            sheet = service.spreadsheets()

            result = (
                sheet.values()
                .get(spreadsheetId=self.sheet_id, range=self.cell_range)
                .execute()
            )
            values = result.get("values", [])

            if not values:
                print("No data found.")
                return

            for row in values:
                print(f"Row: {row}")

        except HttpError as err:
            print(err)


class Docs(Creds):
    def __init__(self, token_path=TOKEN_JSON, scopes=SCOPES, secret=SECRET, document_id=DOCUMENT_ID):
        super().__init__(token_path, scopes, secret)
        self.document_id = document_id

    def get_docs(self):
        self.generated_creds()

        try:
            service = build("docs", "v1", credentials=self.credentials)

            document = service.documents().get(documentId=self.document_id).execute()

            print(f"The title of the document is: {document.get('title')}")

            self.set_docs(str(document))

        except HttpError as err:
            print(err)

    def set_docs(self, text: str):
        self.generated_creds()

        try:
            service = build("docs", "v1", credentials=self.credentials)
            requests = [
                {
                    "insertText": {
                        "text": text,
                        "location": {
                            "index": 1
                        }
                    }
                }
            ]

            service.documents().batchUpdate(documentId=DOCUMENT_ID,
                                            body={'requests': requests}).execute()

            return

        except HttpError as err:
            print(err)


def main():
    sheets = Sheets(cell_range=input(), sheet_id=SHEET_ID)
    sheets.read_sheet()


if __name__ == "__main__":
    main()
