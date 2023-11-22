import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from docx import Document

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/documents.readonly"]

# The ID of a sample document.
DOCUMENT_ID = "1Fh_FkvZDyRyORb2QluqGV1MyKekuBoWUkBfUL94b73w"


def main():
    """Reads paragraphs and structural elements from a Google Docs document
    and writes them to a new .docx file.
    """
    creds = None
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
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("docs", "v1", credentials=creds)

        # Retrieve the document's contents from the Docs service.
        document = service.documents().get(documentId=DOCUMENT_ID).execute()

        paragraphs = []
        structural_elements = []

        # Process the document and extract paragraphs and structural elements.
        for element in document.get("body").get("content"):
            if "paragraph" in element:
                paragraph = element.get("paragraph")
                elements = paragraph.get("elements")
                if elements:
                    text = "".join([elem.get("textRun", {}).get("content", "") for elem in elements])
                    paragraphs.append(text)
            elif "table" in element:
                # Handle tables separately if needed.
                pass
            elif "tableOfContents" in element:
                # Handle table of contents separately if needed.
                pass
            else:
                # Process other types of structural elements if needed.
                pass

        # Create a new .docx document and write paragraphs and structural elements.
        doc = Document()
        for paragraph in paragraphs:
            doc.add_paragraph(paragraph)
        for element in structural_elements:
            # Write structural elements to the document if needed.
            pass

        # Save the .docx file.
        doc.save("output.docx")

        print("Extraction complete. Results saved to output.docx.")
    except HttpError as err:
        print(err)


if __name__ == "__main__":
    main()

#writes the doc file into .docx file, but formattings aren't preserved