import asyncio
import json
import os
from typing import List, Dict, Tuple
import time
from datetime import datetime
import threading
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed
import urllib.parse
from dotenv import load_dotenv
from traceback import format_exc
import streamlit as st


# 🔹 Microsoft Entra (Azure AD) Credentials (SECURE THESE!)
TENANT_ID = os.getenv("TENANT_ID") or st.secrets.get("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID") or st.secrets.get("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET") or st.secrets.get("CLIENT_SECRET")
MAILBOX_ADDRESS = os.getenv("MAILBOX_ADDRESS") or st.secrets.get("MAILBOX_ADDRESS")
BATCH_SIZE = 50  # Fetch emails in batches of 10
MAX_WORKERS = 5  # Number of concurrent threads
MAX_RETRIES = 3  # Max retry attempts
RETRY_DELAY = 10  # Delay between retries (in seconds)

# Airtable Config
AIRTABLE_API_KEY = os.getenv("AIRTABLE_API_KEY") or st.secrets.get("AIRTABLE_API_KEY")
AIRTABLE_BASE_ID = os.getenv("AIRTABLE_BASE_ID") or st.secrets.get("AIRTABLE_BASE_ID")
AIRTABLE_TABLE_NAME = os.getenv("AIRTABLE_TABLE_NAME") or st.secrets.get("AIRTABLE_TABLE_NAME")

GRAPH_API_URL = "https://graph.microsoft.com/v1.0"

HEADERS_AIRTABLE = {
    "Authorization": f"Bearer {AIRTABLE_API_KEY}",
    "Content-Type": "application/json"
}

# Function to format dates for Airtable
def format_date_for_airtable(date_string):
    try:
        dt = datetime.strptime(date_string, "%Y-%m-%dT%H:%M:%SZ")
        return dt.strftime("%Y/%m/%d")
    except ValueError:
        return None

# OAuth2 Token Retrieval
def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    payload = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials",
        "scope": "https://graph.microsoft.com/.default",
    }
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    response = requests.post(url, data=payload, headers=headers)
    return response.json().get("access_token")

# Get total email count in Sent folder
def get_total_email_count(access_token):
    url = f"{GRAPH_API_URL}/users/{MAILBOX_ADDRESS}/mailFolders/SentItems"
    headers = {"Authorization": f"Bearer {access_token}"}
    response = requests.get(url, headers=headers)
    data = response.json()
    return data.get("totalItemCount", 0)

def log_message(message, callback_fn=None):
    """Send log message to logger"""
    if callback_fn:
        callback_fn(message)  # Use the callback function
    print(message)  # Keep console logging
    
def extract_emails(data):
    result = {}
    for recipient in data.get("toRecipients", []) + data.get("ccRecipients", []) + data.get("bccRecipients", []):
        email = recipient["emailAddress"]["address"]
        name = recipient["emailAddress"]["name"]
        result[email] = name
    return result

# Fetch emails from Outlook
def fetch_sent_emails(batch_number, email_data, lock, access_token, callback_fn=None):
    skip = batch_number * BATCH_SIZE
    url = f"{GRAPH_API_URL}/users/{MAILBOX_ADDRESS}/mailFolders/SentItems/messages?$top={BATCH_SIZE}&$skip={skip}&$select=subject,sender,toRecipients,ccRecipients,bccRecipients,receivedDateTime"
    
    headers = {"Authorization": f"Bearer {access_token}"}
    retries = 0
    
    while retries < MAX_RETRIES:
        response = requests.get(url, headers=headers)

        if response.status_code == 200:
            messages = response.json().get("value", [])
            batch_emails = []
            for msg in messages:
                
                email_info = {
                    "subject": msg.get("subject", "No Subject"),
                    "from": msg.get("sender", {}).get("emailAddress", {}).get("address", "Unknown"),
                    "to": [recipient["emailAddress"]["address"] for recipient in msg.get("toRecipients", [])],
                    "cc": [recipient["emailAddress"]["address"] for recipient in msg.get("ccRecipients", [])],
                    "bcc": [recipient["emailAddress"]["address"] for recipient in msg.get("bccRecipients", [])],
                    "received": msg.get("receivedDateTime", "Unknown"),
                    "name_data": extract_emails(msg)
                }
                batch_emails.append(email_info)

            with lock:
                email_data.extend(batch_emails)
                log_message(f"[Thread-{threading.current_thread().name}] ✅ Batch {batch_number} fetched. Total emails so far: {len(email_data)}", callback_fn)
            return
        
        else:
            log_message(f"[Thread-{threading.current_thread().name}] ❌ Error fetching batch {batch_number} (Attempt {retries+1}/{MAX_RETRIES}): {response.text}", callback_fn)
            retries += 1
            time.sleep(RETRY_DELAY)
    
    log_message(f"[Thread-{threading.current_thread().name}] ❌ Batch {batch_number} failed after {MAX_RETRIES} retries.", callback_fn)

# Aggregate email data
def aggregate_email_data(email_data):
    email_stats = {}

    for email in email_data:
        recipients = email["to"] + email["cc"] + email["bcc"]
        for recipient in set(recipients):
            domain = recipient.split("@")[-1] if "@" in recipient else "unknown"

            if recipient not in email_stats:
                email_stats[recipient] = {
                    "Recipient Email": recipient,
                    "Company / Management": domain,
                    "Total Mails Sent": 0,
                    "Name": email.get("name_data", {}).get(recipient, ""),
                    "Last Interacted Date": format_date_for_airtable(email["received"]),
                }

            email_stats[recipient]["Total Mails Sent"] += recipients.count(recipient)
            email_stats[recipient]["Last Interacted Date"] = max(
                email_stats[recipient]["Last Interacted Date"], format_date_for_airtable(email["received"])
            )

    return list(email_stats.values())

# Fetch all existing Airtable records
def fetch_all_airtable_records():
    url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}"
    all_records = []
    params = {"pageSize": 100}  # Max records per request

    while True:
        response = requests.get(url, headers=HEADERS_AIRTABLE, params=params)
        response.raise_for_status()
        data = response.json()
        
        all_records.extend(data.get("records", []))
        
        if "offset" not in data:
            break
        
        params["offset"] = data["offset"]  # Pagination handling

    return all_records

# Push aggregated data to Airtable
def push_to_airtable(aggregated_data, callback_fn=None):
    existing_records = fetch_all_airtable_records()
    existing_records_dict = {
        record["fields"]["Recipient Email"]: {"id": record["id"], "fields": record["fields"]}
        for record in existing_records
    }
    
    log_message(f"Found {len(existing_records)} existing records")

    no_of_updated_records = 0
    no_of_new_records = 0
    update_records = []
    new_records = []

    for entry in aggregated_data:
        email = entry["Recipient Email"].strip().lower()  # Normalize email for comparison
        total_mails_sent = entry["Total Mails Sent"]
        last_interacted = entry["Last Interacted Date"]

        record = existing_records_dict.get(email)  # Fetch record if it exists

        if record:
            existing_fields = record["fields"]
            record_id = record["id"]

            # Check if values are different before updating
            needs_update = (
                existing_fields.get("Total Mails Sent") != total_mails_sent or
                existing_fields.get("Last Interacted Date") != last_interacted.replace("/", "-")
            )

            if needs_update:
                no_of_updated_records += 1
                update_records.append({
                    "id": record_id,
                    "fields": {
                        "Total Mails Sent": total_mails_sent,
                        "Last Interacted Date": last_interacted
                    }
                })
        else:
            no_of_new_records += 1
            new_records.append({"fields": entry})
            
    url = f"https://api.airtable.com/v0/{AIRTABLE_BASE_ID}/{AIRTABLE_TABLE_NAME}"

    # Insert new records
    for i in range(0, len(new_records), 10):
        batch = {"records": new_records[i : i + 10]}
        response = requests.post(url, headers=HEADERS_AIRTABLE, json=batch)
        log_message(f"✅ Uploaded {len(batch['records'])} new records to Airtable", callback_fn)

    # Update existing records
    for i in range(0, len(update_records), 10):
        batch = {"records": update_records[i : i + 10]}
        response = requests.patch(url, headers=HEADERS_AIRTABLE, json=batch)
        log_message(f"🔄 Updated {len(batch['records'])} records in Airtable", callback_fn)
    
    log_message(f"Total no of new records: {no_of_new_records}", callback_fn)
    log_message(f"Total no of updated records: {no_of_updated_records}", callback_fn)    

# Main function
def main(callback_fn=None):
    """Main function that processes emails"""
    def should_stop():
        return hasattr(callback_fn, '__self__') and callback_fn.__self__.should_stop()

    # Check for stop before starting
    if should_stop():
        log_message("🛑 Processing stopped by user", callback_fn)
        return

    access_token = get_access_token()
    if not access_token:
        log_message("❌ Failed to get access token", callback_fn)
        return

    if should_stop():
        log_message("🛑 Processing stopped by user", callback_fn)
        return
    
    total_emails = get_total_email_count(access_token)
    log_message(f"📩 Total emails in Sent folder: {total_emails}", callback_fn)

    total_batches = (total_emails // BATCH_SIZE) + (1 if total_emails % BATCH_SIZE != 0 else 0)
    email_data = []
    lock = threading.Lock()

    def process_batch(batch):
        if should_stop():
            return None
        try:
            # Pass callback_fn to fetch_sent_emails
            return fetch_sent_emails(batch, email_data, lock, access_token, callback_fn)
        except Exception as e:
            log_message(f"❌ Error in batch {batch}: {str(e)}", callback_fn)
            return None

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = []
        for batch in range(total_batches):
            if should_stop():
                log_message("🛑 Processing stopped by user", callback_fn)
                # Cancel any pending futures
                for future in futures:
                    future.cancel()
                return
            futures.append(executor.submit(process_batch, batch))
        
        for future in as_completed(futures):
            if should_stop():
                # Cancel remaining futures
                for f in futures:
                    f.cancel()
                break
            try:
                future.result()
            except Exception as e:
                log_message(f"❌ Error processing batch: {str(e)}", callback_fn)

    if not should_stop():
        # Only process results if not stopped
        try:
            aggregated_data = aggregate_email_data(email_data)
            push_to_airtable(aggregated_data, callback_fn)
        except Exception as e:
            log_message(f"❌ Error in final processing: {str(e)}: {format_exc()}", callback_fn)

if __name__ == "__main__":
    main()
