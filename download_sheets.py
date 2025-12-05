import gspread
import os
import requests
import argparse
from oauth2client.service_account import ServiceAccountCredentials
import time
from tqdm import tqdm

def download_sheet_task(target_sheet, spreadsheet_id, download_dir, headers):
    """Helper function to download a single sheet."""
    try:
        title = target_sheet.title
        gid = target_sheet.id
        
        safe_title = title.replace("/", "-").replace("\\", "-")
        xlsx_path = os.path.join(download_dir, f"{safe_title}.xlsx")
        
        export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx&gid={gid}"
        response = requests.get(export_url, headers=headers)
        
        if response.status_code == 200:
            with open(xlsx_path, 'wb') as f:
                f.write(response.content)
            return True, title
        else:
            return False, f"{title} (Status: {response.status_code})"
    except Exception as e:
        return False, f"{target_sheet.title} (Error: {e})"

def download_sheets(user_input_arg=None):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    
    creds = None
    if "GOOGLE_CREDENTIALS_JSON" in os.environ:
        import json
        creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    else:
        creds_file = 'credentials.json'
        if os.path.exists(creds_file):
            creds = ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)
        else:
            print(f"Error: Credentials not found. Set GOOGLE_CREDENTIALS_JSON env var or place '{creds_file}' in the directory.")
            return

    client = gspread.authorize(creds)

    sheet_url = "https://docs.google.com/spreadsheets/d/1TPdMqE_-MVu2ywkSYyetmJVDtv9CHkjNVkUmYnsBbgE/edit"

    try:
        sheet = client.open_by_url(sheet_url)
        spreadsheet_id = sheet.id
        print(f"Successfully opened spreadsheet: {sheet.title}")

        download_dir = os.path.join("Downloads")
        os.makedirs(download_dir, exist_ok=True)

        all_worksheets = sheet.worksheets()
        excluded_titles = ["KEYWORDS", "Format ST", "Format OP"]
        
        filtered_worksheets = [ws for ws in all_worksheets if ws.title not in excluded_titles]

        if not filtered_worksheets:
            print("No worksheets available to download (all filtered out).")
            return

        print("\nAvailable Worksheets:")
        for i, ws in enumerate(filtered_worksheets):
            print(f"{i + 1}. {ws.title}")

        if user_input_arg:
            user_input = user_input_arg
            print(f"\nUsing input from argument: {user_input}")
        else:
            user_input = input("\nEnter specific sheet number or range (e.g., 1 or 1-5): ").strip()

        selected_indices = []
        try:
            if "-" in user_input:
                start_str, end_str = user_input.split("-")
                start = int(start_str)
                end = int(end_str)
                selected_indices = range(start - 1, end) 
            else:
                index = int(user_input)
                selected_indices = [index - 1]
        except ValueError:
            print("Invalid input format. Please enter a number or a range (e.g., 1-5).")
            return

        access_token = creds.get_access_token().access_token
        headers = {'Authorization': f'Bearer {access_token}'}

        tasks = []
        valid_indices = [idx for idx in selected_indices if 0 <= idx < len(filtered_worksheets)]
        
        if not valid_indices:
            print("No valid sheets selected.")
            return

        print(f"\nStarting download of {len(valid_indices)} sheets...")

        print(f"\nStarting download of {len(valid_indices)} sheets...")

        for idx in tqdm(valid_indices, unit="sheet"):
            target_sheet = filtered_worksheets[idx]
            while True:
                success, result = download_sheet_task(target_sheet, spreadsheet_id, download_dir, headers)
                if success:
                    print(f"Downloaded: {result}")
                    break
                else:
                    print(f"\nFailed to download: {result}. Retrying in 2 seconds...")
                    time.sleep(2)
        
        count = len(valid_indices)
        print(f"\nCompleted. Downloaded {count} files to: {download_dir}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Download Google Sheets worksheets.")
    parser.add_argument("--input", type=str, help="Specific sheet number or range (e.g., 1 or 1-5)")
    args = parser.parse_args()
    
    download_sheets(args.input)
