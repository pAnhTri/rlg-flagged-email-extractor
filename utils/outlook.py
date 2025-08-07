import win32com.client
from datetime import date
import calendar
import os
from utils import get_config

outlook = None
emails = None
main_folder = None
inbox = None


def is_outlook_installed():
    global outlook, main_folder, inbox, emails
    primary_email = get_config("Email", "primary_email")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        main_folder = outlook.Folders(primary_email)

        inbox = main_folder.Folders("Inbox")

        emails = inbox.Items

        return True
    except Exception as e:
        print(f"Error: {e}")
        return False


def get_flagged_emails():
    flagged_emails = []
    for email in emails:
        try:
            if email.FlagStatus == 2:
                flagged_emails.append(email)
        except Exception as e:
            print(f"Error: {e}")
            continue
    return flagged_emails


def get_flagged_emails_in_month(start: date = None, end: date = None):
    flagged_emails = get_flagged_emails()

    start_of_month = date.today().replace(day=1)
    _, num_days = calendar.monthrange(date.today().year, date.today().month)
    end_of_month = date.today().replace(day=num_days)

    flagged_emails_in_month = []

    timeframe_start = start or start_of_month
    timeframe_end = end or end_of_month

    for email in flagged_emails:
        received_date = email.ReceivedTime.date()
        if received_date >= timeframe_start and received_date <= timeframe_end:
            flagged_emails_in_month.append(email)

    return flagged_emails_in_month, start_of_month, end_of_month


def _does_store_exist(store_path):
    for store in outlook.Stores:
        if store.FilePath.lower() == store_path.lower():
            return True
    return False


def get_flagged_emails_in_month_pst(start_of_month, end_of_month):
    flagged_emails_folder_name = f"Flagged Emails {start_of_month.strftime("%m-%d-%y")} - {end_of_month.strftime("%m-%d-%y")}"

    path_from_config = get_config("Folder", "output_folder")
    parsed_path = path_from_config.rstrip("/")

    billing_path = os.path.join(parsed_path, flagged_emails_folder_name) + ".pst"

    # Only create the folder if it doesn't exist
    if not _does_store_exist(billing_path):
        outlook.AddStore(billing_path)
    else:
        print(f"Folder {flagged_emails_folder_name} already exists")

    # Get the store with the path
    flagged_emails_store = None

    for store in outlook.Stores:
        # Normalize paths to handle different separators
        normalized_store_path = os.path.normpath(store.FilePath).lower()
        normalized_billing_path = os.path.normpath(billing_path).lower()
        if normalized_store_path == normalized_billing_path:
            flagged_emails_store = store

    return flagged_emails_store.GetRootFolder()


def copy_flagged_emails_to_pst(flagged_emails_in_month):
    flagged_emails_root = get_flagged_emails_in_month_pst()

    existing_emails_in_store = set()

    for email in flagged_emails_root.Items:
        print(email.EntryID)
        existing_emails_in_store.add(email.EntryID)

    copy_count = 0

    for flagged_email in flagged_emails_in_month:
        if flagged_email.EntryID in existing_emails_in_store:
            print(f"Skipping {flagged_email.Subject}: Already exists!")
            continue

        print(f"Copying {flagged_email.Subject} to {flagged_emails_root.Name}")
        flagged_email.Copy().Move(flagged_emails_root)
        copy_count += 1

    print(f"Copied {copy_count} emails to {flagged_emails_root.Name}")
