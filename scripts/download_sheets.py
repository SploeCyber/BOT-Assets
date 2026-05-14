"""
Download and manage Google Sheets worksheets.

This module handles downloading Google Sheets worksheets with change detection,
hash-based deduplication, and cleanup of deleted worksheets.
"""

import gspread
import os
import shutil
import requests
import argparse
import hashlib
import json
import logging
import time
from typing import Optional, List, Dict, Set
from functools import wraps
from oauth2client.service_account import ServiceAccountCredentials
from tqdm import tqdm

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.dirname(SCRIPT_DIR)

ASSETS_FOLDER = os.path.join(BASE_DIR, "assets")
HASH_FILE = os.path.join(ASSETS_FOLDER, ".sheet-hashes.json")
DOWNLOAD_DIR = os.path.join(BASE_DIR, "Downloads")

# Retry configuration for API Rate Limiting (429 errors)
MAX_RETRIES = 5
RETRY_DELAY = 10  # seconds (increased for 429 errors)
RETRY_BACKOFF = 2  # multiplier


def retry_with_backoff(max_retries: int = MAX_RETRIES, initial_delay: float = RETRY_DELAY, backoff_factor: float = RETRY_BACKOFF):
    """Decorator that retries a function with exponential backoff."""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            delay = initial_delay
            last_exception = None
            
            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_exception = e
                    if attempt < max_retries:
                        logger.warning(f"Attempt {attempt + 1}/{max_retries + 1} failed: {e}. Retrying in {delay}s...")
                        time.sleep(delay)
                        delay *= backoff_factor
                    else:
                        logger.error(f"All {max_retries + 1} attempts failed: {e}")
            
            raise last_exception
        return wrapper
    return decorator


def sanitize_filename(name: str) -> str:
    """Removes non-printable characters and slashes. Matches generate_info_json.py."""
    if not name:
        return ""
    import re
    # Remove non-printable characters (0-31, 127-159)
    name = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', str(name))
    # Replace slashes
    name = name.replace("/", "-").replace("\\", "-")
    return name.strip()


def load_sheet_hashes() -> Dict[str, str]:
    """Load stored per-worksheet content hashes."""
    if os.path.exists(HASH_FILE):
        with open(HASH_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_sheet_hashes(hashes: Dict[str, str]) -> None:
    """Save per-worksheet content hashes."""
    os.makedirs(ASSETS_FOLDER, exist_ok=True)
    with open(HASH_FILE, "w", encoding="utf-8") as f:
        json.dump(hashes, f, ensure_ascii=False, indent=2)


@retry_with_backoff(max_retries=MAX_RETRIES, initial_delay=RETRY_DELAY)
def compute_worksheet_hash(worksheet) -> str:
    """Compute MD5 hash of a worksheet's cell values."""
    all_values = worksheet.get_all_values()
    content = json.dumps(all_values, ensure_ascii=False, sort_keys=True)
    return hashlib.md5(content.encode("utf-8")).hexdigest()


def get_credentials() -> Optional[ServiceAccountCredentials]:
    """Get Google API credentials."""
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

    if "GOOGLE_CREDENTIALS_JSON" in os.environ:
        creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
        return ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)

    creds_file = 'credentials.json'
    if os.path.exists(creds_file):
        return ServiceAccountCredentials.from_json_keyfile_name(creds_file, scope)

    logger.error("Credentials not found. Set GOOGLE_CREDENTIALS_JSON env var or place 'credentials.json' in the directory.")
    return None


@retry_with_backoff(max_retries=MAX_RETRIES, initial_delay=RETRY_DELAY)
def get_filtered_worksheets(client) -> tuple:
    """Open spreadsheet and return filtered worksheets.
    
    Returns:
        tuple: (sheet object, list of filtered worksheets)
    """
    sheet_url = "https://docs.google.com/spreadsheets/d/1TPdMqE_-MVu2ywkSYyetmJVDtv9CHkjNVkUmYnsBbgE/edit"
    sheet = client.open_by_url(sheet_url)

    excluded_titles = ["KEYWORDS", "Format ST", "Format OP"]
    all_worksheets = sheet.worksheets()
    filtered = [ws for ws in all_worksheets if ws.title not in excluded_titles]

    return sheet, filtered


@retry_with_backoff(max_retries=MAX_RETRIES, initial_delay=RETRY_DELAY, backoff_factor=RETRY_BACKOFF)
def download_single_sheet(worksheet, spreadsheet_id: str, headers: Dict[str, str]) -> bool:
    """Download a single worksheet as XLSX with retry logic.
    
    Args:
        worksheet: The worksheet object to download
        spreadsheet_id: The spreadsheet ID
        headers: HTTP headers for authentication
        
    Returns:
        bool: True on success, False on failure
    """
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    # Clear any existing files in download dir
    for f in os.listdir(DOWNLOAD_DIR):
        if f.endswith(".xlsx"):
            os.remove(os.path.join(DOWNLOAD_DIR, f))

    title = worksheet.title
    gid = worksheet.id
    safe_title = sanitize_filename(title)
    xlsx_path = os.path.join(DOWNLOAD_DIR, f"{safe_title}.xlsx")

    export_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=xlsx&gid={gid}"

    logger.info(f"Downloading: {title}")
    response = requests.get(export_url, headers=headers, timeout=60)
    response.raise_for_status()
    
    with open(xlsx_path, 'wb') as f:
        f.write(response.content)
    
    logger.info(f"✅ Downloaded: {xlsx_path}")
    return True


def cmd_check(args) -> None:
    """Check which worksheets have changed. Outputs changed indices (1-based)."""
    creds = get_credentials()
    if not creds:
        return

    client = gspread.authorize(creds)
    sheet, filtered = get_filtered_worksheets(client)
    stored_hashes = load_sheet_hashes()

    changed = []
    hashes_to_update = {}

    for i, ws in enumerate(tqdm(filtered, unit="sheet", desc="Checking", mininterval=5)):
        try:
            safe_title = sanitize_filename(ws.title)
            folder_path = os.path.join(ASSETS_FOLDER, safe_title)
            folder_exists = os.path.exists(folder_path)
            
            stored_hash = stored_hashes.get(ws.title)
            
            # Fast check: skip hashing if we have a hash and the folder exists
            if args.skip_hash and stored_hash and folder_exists:
                continue

            current_hash = compute_worksheet_hash(ws)

            is_changed = False
            if args.force:
                is_changed = True
                status = "FORCE"
            elif not stored_hash:
                is_changed = True
                # Use different status to clarify why it's flagged
                status = "UNTRACKED (LOCAL EXISTS)" if folder_exists else "NEW"
            elif current_hash != stored_hash:
                is_changed = True
                status = "CHANGED"
            
            if is_changed:
                changed.append(i + 1)  # 1-based
                logger.info(f"  {status}: {ws.title}")
                if args.update_hashes:
                    hashes_to_update[ws.title] = current_hash
            
            time.sleep(1)  # Rate limiting
        except Exception as e:
            logger.error(f"  Error checking {ws.title}: {e}. Marking as changed.")
            changed.append(i + 1)

    if args.update_hashes and hashes_to_update:
        stored_hashes.update(hashes_to_update)
        save_sheet_hashes(stored_hashes)
        logger.info(f"✅ Updated {len(hashes_to_update)} hashes in {HASH_FILE}")

    if changed:
        # Output changed indices as space-separated for shell consumption
        logger.info(f"\nCHANGED_INDICES={' '.join(str(i) for i in changed)}")
    else:
        logger.info("\nCHANGED_INDICES=")
        logger.info("No sheets changed.")


def cmd_download_one(args) -> None:
    """Download a single worksheet by index, with optional hash check."""
    creds = get_credentials()
    if not creds:
        return

    client = gspread.authorize(creds)
    sheet, filtered = get_filtered_worksheets(client)

    idx = args.index - 1  # Convert to 0-based
    if idx < 0 or idx >= len(filtered):
        logger.error(f"Invalid index {args.index}. Valid range: 1-{len(filtered)}")
        return

    ws = filtered[idx]
    stored_hashes = load_sheet_hashes()

    # Check hash unless forced
    if not args.force:
        try:
            current_hash = compute_worksheet_hash(ws)
            stored_hash = stored_hashes.get(ws.title)

            if current_hash == stored_hash:
                logger.info(f"⏭️  SKIPPED: {ws.title} (unchanged)")
                return

            # Update hash
            stored_hashes[ws.title] = current_hash
        except Exception as e:
            logger.error(f"  Hash check error: {e}. Downloading anyway.")
            current_hash = None

    # Download
    logger.info(f"⬇️  Downloading: {ws.title}")
    access_token = creds.get_access_token().access_token
    headers = {'Authorization': f'Bearer {access_token}'}

    try:
        success = download_single_sheet(ws, sheet.id, headers)

        if success:
            # Save updated hash
            if not args.force:
                save_sheet_hashes(stored_hashes)
            logger.info(f"✅ DOWNLOADED: {ws.title}")
        else:
            logger.error(f"❌ FAILED: {ws.title}")
    except Exception as e:
        logger.error(f"❌ Download failed for {ws.title}: {e}")


def cmd_cleanup(args) -> None:
    """Remove assets folders and hashes for worksheets that no longer exist in the Google Sheet."""
    creds = get_credentials()
    if not creds:
        return

    client = gspread.authorize(creds)
    sheet, filtered = get_filtered_worksheets(client)

    # Current worksheet titles from Google Sheets
    current_titles = set(ws.title for ws in filtered)
    # Also build sanitized variants used for folder names
    current_safe_titles = set(sanitize_filename(title) for title in current_titles)

    stored_hashes = load_sheet_hashes()
    deleted_any = False

    # 1. Check assets/ folders
    if os.path.exists(ASSETS_FOLDER):
        for folder_name in sorted(os.listdir(ASSETS_FOLDER)):
            folder_path = os.path.join(ASSETS_FOLDER, folder_name)
            if not os.path.isdir(folder_path):
                continue
            # Check if this folder matches any current worksheet
            if folder_name not in current_titles and folder_name not in current_safe_titles:
                logger.info(f"🗑️  DELETED: {folder_name}")
                shutil.rmtree(folder_path)
                deleted_any = True

    # 2. Clean orphaned hash entries
    orphaned_keys = [key for key in stored_hashes if key not in current_titles]
    for key in orphaned_keys:
        logger.info(f"🗑️  HASH_REMOVED: {key}")
        del stored_hashes[key]
        deleted_any = True

    if deleted_any:
        save_sheet_hashes(stored_hashes)
        # Regenerate asset-map.json
        _regenerate_asset_map()
        logger.info("\n✅ Cleanup complete.")
    else:
        logger.info("✅ No deleted worksheets found.")


def _regenerate_asset_map() -> None:
    """Regenerate asset-map.json from current assets/ folders."""
    asset_list = []
    
    # Add all-cards.json as the first entry (merged dataset)
    all_cards_path = os.path.join(ASSETS_FOLDER, "all-cards.json")
    if os.path.exists(all_cards_path):
        asset_list.append({
            "name": "All Cards",
            "path": "assets/all-cards.json"
        })
    
    if os.path.exists(ASSETS_FOLDER):
        for folder_name in sorted(os.listdir(ASSETS_FOLDER)):
            folder_path = os.path.join(ASSETS_FOLDER, folder_name)
            if os.path.isdir(folder_path):
                dataset_path = os.path.join(folder_path, "card-data.json")
                if os.path.exists(dataset_path):
                    asset_list.append({
                        "name": folder_name,
                        "path": f"assets/{folder_name}/card-data.json"
                    })
    with open(os.path.join(BASE_DIR, "asset-map.json"), "w", encoding="utf-8") as f:
        json.dump(asset_list, f, ensure_ascii=False, indent=2)
    logger.info(f"📦 Asset map updated ({len(asset_list)} entries)")


def cmd_count(args) -> None:
    """Output the total number of worksheets."""
    creds = get_credentials()
    if not creds:
        return

    client = gspread.authorize(creds)
    sheet, filtered = get_filtered_worksheets(client)
    print(len(filtered))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Download Google Sheets worksheets.")
    subparsers = parser.add_subparsers(dest="command", help="Commands")
    
    # count: output total worksheet count
    sub_count = subparsers.add_parser("count", help="Output total worksheet count")
    sub_count.set_defaults(func=cmd_count)
    
    # check: check which worksheets changed
    sub_check = subparsers.add_parser("check", help="Check which worksheets have changed")
    sub_check.add_argument("--force", action="store_true", help="Mark all as changed")
    sub_check.add_argument("--update-hashes", action="store_true", help="Update hash file with current sheet content without downloading")
    sub_check.add_argument("--skip-hash", action="store_true", help="Skip hashing if folder exists and hash is tracked (fast check)")
    sub_check.set_defaults(func=cmd_check)
    
    # download-one: download a single worksheet
    sub_one = subparsers.add_parser("download-one", help="Download a single worksheet by index (1-based)")
    sub_one.add_argument("index", type=int, help="Worksheet index (1-based)")
    sub_one.add_argument("--force", action="store_true", help="Skip hash check")
    sub_one.set_defaults(func=cmd_download_one)
    
    # cleanup: remove data for deleted worksheets
    sub_cleanup = subparsers.add_parser("cleanup", help="Remove assets folders for deleted worksheets")
    sub_cleanup.set_defaults(func=cmd_cleanup)
    
    args = parser.parse_args()
    
    if hasattr(args, 'func'):
        args.func(args)
    else:
        parser.print_help()
