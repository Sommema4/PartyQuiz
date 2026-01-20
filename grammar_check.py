"""
Grammar Check Script for PartyQuiz
This script reads questions from Google Sheets, checks their grammar using LLM,
and writes the corrected versions to column D.
"""

import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from llm_interface import LLMInterface

# Google Spreadsheet details
SPREADSHEET_ID = "19o12YCNMoeEY4UBvScqW2vDv35rbC21RelnUcn__ExM"
SHEET_NAME = 'otazky'
READ_RANGE = 'otazky!A:C'  # Read columns A-C (Téma, Otázka, Odpověď)

# Scopes - need both read and write for spreadsheets
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
]


def get_credentials():
    """Handles OAuth2 authentication and returns credentials."""
    creds = None
    token_file = 'token_grammar.json'  # Separate token file for grammar checker
    
    # Token file stores the user's access and refresh tokens
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)
    
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
        with open(token_file, 'w') as token:
            token.write(creds.to_json())
    
    return creds


def read_questions_from_sheet(creds, spreadsheet_id, range_name):
    """Reads questions from Google Spreadsheet."""
    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])
        
        if not values:
            print('❌ No data found in spreadsheet.')
            return []
        
        print(f'✓ Found {len(values)} rows of data')
        return values
    except HttpError as error:
        print(f"❌ An error occurred reading spreadsheet: {error}")
        return []


def write_corrections_to_sheet(creds, spreadsheet_id, sheet_name, corrections):
    """
    Writes corrected questions to column D of the spreadsheet.
    
    Args:
        corrections: List of dicts with 'row' (1-indexed), 'original', 'corrected'
    """
    if not corrections:
        print("ℹ️  No corrections to write.")
        return
    
    try:
        service = build('sheets', 'v4', credentials=creds)
        
        # Prepare batch update data
        data = []
        for correction in corrections:
            row_num = correction['row']
            corrected_text = correction['corrected']
            
            data.append({
                'range': f'{sheet_name}!D{row_num}',
                'values': [[corrected_text]]
            })
        
        # Batch update
        body = {
            'valueInputOption': 'RAW',
            'data': data
        }
        
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body=body
        ).execute()
        
        updated_cells = result.get('totalUpdatedCells', 0)
        print(f"✓ Updated {updated_cells} cells in column D")
        
    except HttpError as error:
        print(f"❌ An error occurred writing to spreadsheet: {error}")


def check_grammar_and_collect_corrections(data, llm):
    """
    Checks grammar for all questions and collects corrections.
    
    Returns:
        List of dicts with 'row', 'original', 'corrected', 'confidence', 'has_errors'
    """
    print("\n" + "="*70)
    print("🤖 Starting Grammar Check")
    print("="*70 + "\n")
    
    corrections = []
    
    # Skip header row (row 1)
    for idx, row in enumerate(data[1:], start=2):  # Start from row 2 (1-indexed)
        if not row or len(row) == 0 or not row[0].strip():
            continue
        
        question_text = row[0].strip()
        print(f"Row {idx}: Checking '{question_text[:60]}...'")
        
        # Check grammar
        grammar_result = llm.check_czech_grammar(question_text)
        
        if grammar_result.get('has_errors'):
            confidence = grammar_result.get('confidence', 0)
            print(f"  ⚠️  Grammar issues found (confidence: {confidence}%)")
            
            # Get the corrected text
            corrections_list = grammar_result.get('corrections', [])
            if corrections_list:
                # Apply all corrections
                corrected_text = question_text
                for correction_item in corrections_list:
                    original = correction_item.get('original', '')
                    corrected = correction_item.get('corrected', '')
                    if original and corrected:
                        corrected_text = corrected_text.replace(original, corrected)
                        print(f"     • '{original}' → '{corrected}'")
                
                corrections.append({
                    'row': idx,
                    'original': question_text,
                    'corrected': corrected_text,
                    'confidence': confidence,
                    'has_errors': True
                })
            else:
                # LLM says there are errors but didn't provide corrections
                # Write a note in column D
                corrections.append({
                    'row': idx,
                    'original': question_text,
                    'corrected': f"[LLM detekoval chyby ale neposkytl opravu - jistota: {confidence}%]",
                    'confidence': confidence,
                    'has_errors': True
                })
        else:
            confidence = grammar_result.get('confidence', 0)
            print(f"  ✓ No issues (confidence: {confidence}%)")
            # Don't add to corrections list - leave column D blank
        
        if 'error' in grammar_result:
            print(f"  ⚠️  Error: {grammar_result['error']}")
    
    print("\n" + "="*70)
    print(f"✓ Grammar check complete")
    print(f"  Total questions checked: {len(data) - 1}")  # Minus header row
    print(f"  Questions with errors: {len(corrections)}")
    print("="*70 + "\n")
    
    return corrections


def main():
    """Main function to run the grammar check."""
    print("\n" + "="*70)
    print("PartyQuiz - Grammar Checker")
    print("="*70 + "\n")
    
    # Initialize LLM
    print("🔧 Initializing LLM interface...")
    llm = LLMInterface()
    print("✓ LLM interface ready\n")
    
    # Get credentials
    print("🔑 Authenticating with Google...")
    creds = get_credentials()
    print("✓ Authentication successful\n")
    
    # Read data from spreadsheet
    print("📊 Reading data from spreadsheet...")
    data = read_questions_from_sheet(creds, SPREADSHEET_ID, READ_RANGE)
    
    if not data or len(data) < 2:
        print("❌ No questions found to check.")
        return
    
    # Check grammar and collect corrections
    corrections = check_grammar_and_collect_corrections(data, llm)
    
    # Write corrections to column D
    print("📝 Writing corrections to spreadsheet (column D)...")
    write_corrections_to_sheet(creds, SPREADSHEET_ID, SHEET_NAME, corrections)
    
    print("\n" + "="*70)
    print("✓ Grammar check complete!")
    print("  • Corrected questions saved to column D")
    print("="*70 + "\n")


if __name__ == "__main__":
    main()
