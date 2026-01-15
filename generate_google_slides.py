import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Google Spreadsheet and Slides details
SPREADSHEET_ID = "19o12YCNMoeEY4UBvScqW2vDv35rbC21RelnUcn__ExM"
RANGE_NAME = 'otazky!A:C'  # All rows, columns A-C (Téma, Odpovědi, Poznámky)

# Scopes - what permissions we need
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/presentations',
    'https://www.googleapis.com/auth/drive'
]

def get_credentials():
    """Handles OAuth2 authentication and returns credentials."""
    creds = None
    
    # Token file stores the user's access and refresh tokens
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # If there are no (valid) credentials available, let the user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired credentials...")
            creds.refresh(Request())
        else:
            print("Opening browser for authentication...")
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds


def read_spreadsheet(creds, spreadsheet_id, range_name):
    """Reads data from a Google Spreadsheet."""
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])
        
        if not values:
            print('No data found.')
            return []
        
        print(f'Found {len(values)} rows of data')
        return values
    except HttpError as error:
        print(f"An error occurred reading spreadsheet: {error}")
        return []


def parse_quiz(data):
    """
    Parses quiz data from spreadsheet format.
    Returns list of topics, where each topic contains:
    - topic_name: Name from "Téma" column
    - questions: List of dicts with 'question', 'answer', 'notes'
    """
    if not data or len(data) < 2:
        return []
    
    topics = []
    current_topic = None
    i = 1  # Skip header row
    
    while i < len(data):
        row = data[i]
        
        # Check if row is empty (or has empty first column)
        if not row or not row[0].strip():
            i += 1
            continue
        
        # Check if this is a topic row (has value in "Téma" column but next rows are questions)
        # A topic row is followed by question rows
        is_topic = False
        if i + 1 < len(data):
            next_row = data[i + 1]
            # If next row exists and has content, this might be a topic
            if next_row and len(next_row) > 0:
                is_topic = True
        
        if is_topic and (current_topic is None or (current_topic and len(current_topic['questions']) >= 6)):
            # Start new topic
            current_topic = {
                'topic_name': row[0].strip(),
                'questions': []
            }
            topics.append(current_topic)
            i += 1
            continue
        
        # This is a question row
        if current_topic is not None:
            question = {
                'question': row[0].strip() if len(row) > 0 else '',
                'answer': row[1].strip() if len(row) > 1 else '',
                'notes': row[2].strip() if len(row) > 2 else ''
            }
            if question['question']:  # Only add if question exists
                current_topic['questions'].append(question)
        
        i += 1
    
    return topics


def create_presentation(creds, title, topics_data, folder_id=None):
    """
    Creates a Google Slides presentation with quiz questions.
    Each slide has:
    - Title: Topic name
    - Body: Question
    """
    try:
        service = build('slides', 'v1', credentials=creds)
        
        # Create a new presentation
        body = {'title': title}
        presentation = service.presentations().create(body=body).execute()
        presentation_id = presentation.get('presentationId')
        
        # Move to the same folder as the spreadsheet if folder_id provided
        if folder_id:
            try:
                drive_service = build('drive', 'v3', credentials=creds)
                file = drive_service.files().get(fileId=presentation_id, fields='parents').execute()
                previous_parents = ",".join(file.get('parents', []))
                
                print(f"  Moving from: {previous_parents}")
                print(f"  Moving to: {folder_id}")
                
                # Update file to move to the shared folder
                result = drive_service.files().update(
                    fileId=presentation_id,
                    addParents=folder_id,
                    removeParents=previous_parents,
                    fields='id, parents'
                ).execute()
                print(f"  ✓ Moved successfully! New parents: {result.get('parents', [])}")
            except Exception as e:
                print(f"  Warning: Could not move presentation: {e}")
                print(f"  Presentation remains in My Drive")
        
        print(f"Created presentation: {title}")
        print(f"ID: {presentation_id}")
        print(f"URL: https://docs.google.com/presentation/d/{presentation_id}")
        
        # Build requests to add slides and delete the default first slide
        requests = []
        
        # First, delete the default title slide that Google creates
        default_slide_id = presentation.get('slides', [{}])[0].get('objectId')
        if default_slide_id:
            requests.append({
                'deleteObject': {
                    'objectId': default_slide_id
                }
            })
        
        for topic in topics_data:
            topic_name = topic['topic_name']
            for question in topic['questions']:
        # Add all question slides
                # Create a slide for each question
                requests.append({
                    'createSlide': {
                        'slideLayoutReference': {
                            'predefinedLayout': 'TITLE_AND_BODY'
                        }
                    }
                })
        
        # Execute all slide creation requests
        if requests:
            body = {'requests': requests}
            response = service.presentations().batchUpdate(
                presentationId=presentation_id, body=body).execute()
            print(f"Added {len(requests)} slides to presentation")
        
        # Now add text content to the slides
        add_content_to_slides(service, presentation_id, topics_data)
        
        return presentation_id
        
    except HttpError as error:
        print(f"An error occurred creating presentation: {error}")
        return None


def add_content_to_slides(service, presentation_id, topics_data):
    """
    Adds text content to slides.
    Each slide: Title = question number + topic name, Body = question without number
    """
    import re
    
    try:
        # Get the presentation to find slide IDs
        presentation = service.presentations().get(
            presentationId=presentation_id).execute()
        slides = presentation.get('slides', [])
        
        requests = []
        slide_idx = 0
        
        for topic in topics_data:
            topic_name = topic['topic_name']
            
            # Extract the number from topic name if it exists (e.g., "1. Aktuality" -> "Aktuality")
            topic_name_clean = re.sub(r'^\d+[\.\)]\s*', '', topic_name).strip()
            
            for question in topic['questions']:
                if slide_idx >= len(slides):
                    break
                
                # Extract number or "T" from question (e.g., "2) how are you?" or "T) bonus question")
                question_text = question['question']
                number_match = re.match(r'^([T\d]+[\.\)])\s*', question_text)
                
                if number_match:
                    question_number = number_match.group(1)
                    question_without_number = question_text[len(number_match.group(0)):].strip()
                    slide_title = f"{question_number} {topic_name_clean}"
                else:
                    # If no number found, just use the full question text and topic name
                    question_without_number = question_text
                    slide_title = topic_name_clean
                
                # Find text boxes on the slide
                page_elements = slides[slide_idx].get('pageElements', [])
                
                if len(page_elements) >= 2:
                    # Title = Question number + Topic name
                    requests.append({
                        'insertText': {
                            'objectId': page_elements[0]['objectId'],
                            'text': slide_title,
                            'insertionIndex': 0
                        }
                    })
                    
                    # Body = Question without number
                    requests.append({
                        'insertText': {
                            'objectId': page_elements[1]['objectId'],
                            'text': question_without_number,
                            'insertionIndex': 0
                        }
                    })
                
                slide_idx += 1
        
        # Execute all text insertion requests
        if requests:
            body = {'requests': requests}
            service.presentations().batchUpdate(
                presentationId=presentation_id, body=body).execute()
            print(f"Added content to {slide_idx} slides")
            
    except HttpError as error:
        print(f"An error occurred adding content: {error}")


def main():
    """Main function to run the script."""
    print("=== Google Sheets to Slides Converter ===\n")
    
    # Authenticate
    creds = get_credentials()
    print("✓ Authentication successful\n")
    
    # Read spreadsheet
    print(f"Reading spreadsheet: {SPREADSHEET_ID}")
    print(f"Range: {RANGE_NAME}")
    data = read_spreadsheet(creds, SPREADSHEET_ID, RANGE_NAME)
    
    if not data:
        print("No data to process. Exiting.")
        return
    
    # Get the folder of the spreadsheet
    folder_id = None
    try:
        drive_service = build('drive', 'v3', credentials=creds)
        file = drive_service.files().get(fileId=SPREADSHEET_ID, fields='parents').execute()
        parents = file.get('parents', [])
        if parents:
            folder_id = parents[0]
            print(f"✓ Found spreadsheet folder ID: {folder_id}")
        else:
            print("Note: Spreadsheet has no parent folder (might be in root)")
    except Exception as e:
        print(f"Note: Could not get spreadsheet folder: {e}")
        print("Presentations will be saved to 'My Drive'")
    
    # Parse quiz data
    topics = parse_quiz(data)
    print(f"\n✓ Parsed {len(topics)} topics")
    
    for i, topic in enumerate(topics):
        print(f"  Topic {i+1}: {topic['topic_name']} ({len(topic['questions'])} questions)")
    
    # Create presentations - 2 topics per presentation
    presentation_ids = []
    for i in range(0, len(topics), 2):
        # Take 2 topics at a time
        batch_topics = topics[i:i+2]
        
        # Create title for presentation
        topic_names = [t['topic_name'] for t in batch_topics]
        title = f"Party Quiz - {' & '.join(topic_names)}"
        
        print(f"\nCreating presentation {i//2 + 1}: {title}")
        presentation_id = create_presentation(creds, title, batch_topics, folder_id)
        
        if presentation_id:
            presentation_ids.append(presentation_id)
    
    # Summary
    print(f"\n{'='*60}")
    print(f"✓ Success! Created {len(presentation_ids)} presentation(s):")
    for i, pid in enumerate(presentation_ids):
        print(f"\n  Presentation {i+1}:")
        print(f"  https://docs.google.com/presentation/d/{pid}")


if __name__ == '__main__':
    main()
