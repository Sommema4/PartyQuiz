import requests
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# Google Spreadsheet and Slides details
SPREADSHEET_ID = "1zFjlmAF7oBPbnfDVmFPs9itCdFem4XTQaBSYM_JgfxM"
RANGE_NAME = 'otazky!A1:C100'
SLIDE_TEMPLATE_ID = ""  # Leave empty if creating a new presentation
API_KEY = "AIzaSyBJMPWrbiHbiyzf0KZ50NrmLmsQVMGjDvg"

def main():
    # Read Google sheet
    sheets = authenticate_sheets(API_KEY)
    quiz_sheet = read_spreadsheet(sheets, SPREADSHEET_ID, RANGE_NAME)

    create_new_presentation('test')

    # quiz = parse_quiz(quiz_sheet)
    # for key, val in quiz.items():
    #     create_slide_request(key, val)
    #     # Execute the requests
    #     # presentation_id = presentation_service['presentations']
    #     presentation_service.presentations().batchUpdate(presentationId='qwd', body={"requests": requests}).execute()
    #     print(f'Presentation created: https://docs.google.com/presentation/d/{presentation_id}')

    # Create the Slides API service
    # presentation_service = authenticate_presentation(API_KEY)


def authenticate_sheets(api_key):
    return build('sheets', 'v4', developerKey=api_key).spreadsheets()

def authenticate_presentation(api_key):
    return build('slides', 'v1', developerKey=API_KEY)

def read_spreadsheet(sheets, spreadsheet_id, range_name):
    result = sheets.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    return values

def parse_quiz(data):
    questions = []
    topics = [line[0] for line in data[::10]]
    for i in range(1, len(data), 10):
        questions.append(data[i:i+6])
    return dict(zip(topics, questions))

def create_slide_request(topic, questions):
    requests = [
    {
        "createSlide": {
            "objectId": "title_slide",
            "insertionIndex": "1",
            "slideLayoutReference": {"predefinedLayout": "TITLE"}
        }
    },
    {
        "insertText": {
            "objectId": "title_slide",
            "text": "Welcome to My Presentation"
        }
    },
    {
        "createSlide": {
            "objectId": "content_slide",
            "insertionIndex": "2",
            "slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"}
        }
    },
    {
        "insertText": {
            "objectId": "content_slide",
            "text": "This is a sample content slide."
        }
    }
    ]
    return requests

def create_new_presentation(title):
    try:
        service = build("slides", "v1", developerKey=API_KEY)

        body = {"title": title}
        presentation = service.presentations().create(body=body).execute()
        print(
            f"Created presentation with ID:{(presentation.get('presentationId'))}"
        )
        return presentation
    except HttpError as error:
        print(f"An error occurred: {error}")
        print("presentation not created")
        return error

if __name__ == '__main__':
    main()