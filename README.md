# High-Level Documentation: Outlook Email Processing & Airtable Sync

## 📌 Overview
This system automates fetching sent emails from Microsoft Outlook using Microsoft Graph API, processes the data, and syncs it with an Airtable database. The solution is designed for scalability and reliability using multithreading.

## 🏗️ Architecture
1. **Authentication**  
   - Uses OAuth2 client credentials flow to get an access token from Microsoft Entra (Azure AD).  
   - Securely retrieves credentials from environment variables or Streamlit secrets.

2. **Fetching Sent Emails**  
   - Retrieves emails from the "Sent Items" folder of the specified mailbox.  
   - Extracts `toRecipients`, `ccRecipients`, and `bccRecipients` along with metadata.  
   - Implements batching (`BATCH_SIZE`) to manage API rate limits.  
   - Uses multithreading (`ThreadPoolExecutor`) to fetch emails concurrently.  
   - Retries failed requests up to `MAX_RETRIES` with exponential backoff.

3. **Data Aggregation**  
   - Groups email records by recipient.  
   - Aggregates `Total Mails Sent` and `Last Interacted Date`.  
   - Extracts recipient domains for classification.

4. **Airtable Sync**  
   - Fetches existing records from Airtable to avoid duplicates.  
   - Updates existing records if email stats have changed.  
   - Inserts new records in batches for efficiency.  

## ⚙️ Configuration  
| Variable Name           | Description                          |
|-------------------------|--------------------------------------|
| `TENANT_ID`            | Microsoft Entra Tenant ID           |
| `CLIENT_ID`            | Azure AD Application Client ID      |
| `CLIENT_SECRET`        | Azure AD Client Secret              |
| `MAILBOX_ADDRESS`      | Email address of the monitored mailbox |
| `BATCH_SIZE`           | Number of emails fetched per request |
| `MAX_WORKERS`          | Number of concurrent threads        |
| `MAX_RETRIES`          | Maximum retries on failed API calls |
| `AIRTABLE_API_KEY`     | Airtable API Key                    |
| `AIRTABLE_BASE_ID`     | Airtable Base ID                    |
| `AIRTABLE_TABLE_NAME`  | Name of the Airtable table          |

## 🔄 Execution Flow  
1. Retrieve an access token from Azure AD.  
2. Fetch total email count from the Sent folder.  
3. Process emails in parallel batches using multithreading.  
4. Aggregate recipient-based statistics.  
5. Fetch existing Airtable records.  
6. Compare and update records if needed.  
7. Insert new records into Airtable.  

## 🚀 Performance Optimizations  
- Uses **ThreadPoolExecutor** for concurrent API requests.  
- Implements **retry logic** with exponential backoff.  
- Uses **batch processing** for efficient API interactions.  
- **Data deduplication** prevents redundant Airtable updates.  

## 🛠️ Error Handling & Logging  
- Implements structured logging with timestamps.  
- Uses a `callback_fn` to allow UI integration for real-time status updates.  
- Gracefully handles API failures, rate limits, and unexpected errors.  

## 📌 Future Improvements  
- Implement caching to reduce duplicate API calls.  
- Add logging to a persistent storage backend.  
- Optimize error handling with retry queues.  
- Implement webhook triggers for real-time email processing.

---
**Author**: Surya kiran
**Last Updated**: `March 2025`  
