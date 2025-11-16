
# Syncro MSP → Zoho Books Invoice Sync

This PowerShell script syncs **unpaid invoices** from **Syncro MSP** into **Zoho Books** and (optionally) marks those invoices as *Quick Paid* in Syncro once they are successfully created in Zoho.

It is intended to be run on a Windows machine (manually or via Task Scheduler) and uses:

- The **Syncro MSP REST API** (API key + subdomain).
- The **Zoho Books OAuth 2.0 API** (authorization code + refresh tokens).
- A simple **CSV-based product mapping** between Syncro products and Zoho Books items.

> By default, the script:
> - Looks up customers that have a `ZohoCustomerId` property in Syncro.
> - Retrieves unpaid invoices since the **first day of the current month**.
> - Maps line items using `ProductMappings.csv`.
> - Creates matching invoices in Zoho Books.
> - Optionally creates a “Quick” payment in Syncro to mark those invoices as paid.

---

## Features

- Pulls all **unpaid Syncro invoices since the first of the month**.
- Only processes customers that have a **Zoho customer ID** stored in a custom property (`ZohoCustomerId`).
- Uses a **CSV file** to map Syncro `product_id` values to Zoho `item_id` values.
- Supports a **default product mapping** for any unmapped items.
- Handles **Zoho OAuth 2.0**:
  - Interactive browser authorization on first run.
  - Saves and reuses **access** and **refresh tokens** in `config.json`.
  - Auto-refreshes access tokens when close to expiry.
- Optionally creates **Quick payments in Syncro** to mark invoices as paid.
- Detailed logging to a daily log file in the script directory.

---

## Prerequisites

### 1. Environment

- Windows with **PowerShell 5.1+** or **PowerShell 7+**.
- Outbound HTTPS access to:
  - `https://accounts.zoho.com` (or your Zoho accounts domain)
  - `https://www.zohoapis.com` (Zoho Books API, US region)
  - `https://<your-subdomain>.syncromsp.com` (Syncro MSP API)

### 2. Syncro MSP

- A Syncro MSP account.
- A **Syncro API Key** with permissions to:
  - Read customers.
  - Read invoices and line items.
  - Create payments.
- Your Syncro **subdomain** (the `<your-subdomain>` portion of `https://<your-subdomain>.syncromsp.com`).
- A customer-level property named **`ZohoCustomerId`** (or equivalent) populated with the corresponding Zoho Books customer ID for each customer you want synced.

### 3. Zoho Books

- A **Zoho account** with **Zoho Books** enabled.
- A Zoho **client** configured for OAuth 2.0 (in Zoho API Console):
  - Client ID
  - Client Secret
  - Redirect URI (must match what you put in `config.json`).
- Your **Zoho Books Organization ID** (Org ID).
- API scopes that include access to Zoho Books (for example: `ZohoBooks.fullaccess.all`, adjust as needed).

> Note: This script currently uses **US-region endpoints** for Zoho Books  
> (`https://www.zohoapis.com/books/v3/...`). If your Zoho Books organization is in another region (EU, IN, etc.), you will need to adjust the API base URL in the script.

---

## Repository Files

Typical layout:

- `REAL-SyncroToBooks.ps1`  
  The main script.

- `config.example.json`  
  Example configuration file. Copy to `config.json` and fill in your credentials.

- `config.json`  
  Your actual configuration (not tracked in git; should be excluded via `.gitignore`).

- `ProductMappings.csv`  
  CSV file defining how Syncro product IDs map to Zoho Books item IDs.

- `SyncroToBooks-Log-YYYYMMDD.txt`  
  Daily log files created by the script (one per run/day).

---

## Configuration

### 1. Create `config.json`

Copy `config.example.json` to `config.json` in the same folder as the script and fill in your values.

Example structure:

```json
{
  "Zoho": {
    "ClientID": "1000.xxxxxx",
    "Secret": "your-zoho-client-secret",
    "RedirectUri": "https://localhost",
    "AuthorizeUri": "https://accounts.zoho.com/oauth/v2/auth",
    "Scope": "ZohoBooks.fullaccess.all",
    "OrganizationID": "1234567890123",

    "AccessToken": "",
    "RefreshToken": "",
    "TokenExpiration": ""
  },
  "Syncro": {
    "APIKey": "your-syncro-api-key",
    "Subdomain": "your-syncro-subdomain"
  }
}

The script will:

- Validate that all required Zoho and Syncro fields exist.
- On first run, open a browser for Zoho OAuth, obtain tokens, and **write `AccessToken`, `RefreshToken`, and `TokenExpiration` back into `config.json`** using the `Save-ZohoTokens` function.

> Treat `config.json` as sensitive. It contains API credentials and tokens.  
> Restrict file permissions and do not commit it to a public repository.

### 2. ProductMappings.csv

The script expects a `ProductMappings.csv` file in the same directory, with at least these columns:

- `SyncroProductID`  
- `ZohoItemID`  
- `ProductName`  
- `IncludeDescription` (`Yes` / `No`)

Example:

```csv
SyncroProductID,ZohoItemID,ProductName,IncludeDescription
DEFAULT,1234567890001,Default Service,Yes
42,1234567890002,Monthly Managed Service,Yes
99,1234567890003,Hardware Sale,No
```

Behavior:

- If a line item’s `product_id` matches a `SyncroProductID`, that mapping is used.
- If no mapping is found, the `DEFAULT` row is used (if defined).
- If `IncludeDescription` is `"Yes"`, the Syncro line item name is included as the description in Zoho.

The script validates that:

- `ProductMappings.csv` exists.
- All required columns are present.
- A lookup table is built in memory for quick mapping.

---

## How It Works (High-Level Flow)

1. **Startup & Logging**
   - Creates a log file named `SyncroToBooks-Log-YYYYMMDD.txt` in the script directory.
   - Logs start time and whether `-SkipQuickPay` was specified.

2. **Config & Token Handling**
   - Loads `config.json` and validates required fields.
   - Reads Zoho and Syncro settings into variables.
   - If a refresh token and token expiration are present:
     - Checks whether the access token is still valid.
     - If expiring or expired, calls `Refresh-ZohoAccessToken`.
   - If no valid refresh token is available:
     - Constructs an authorization URL using `AuthorizeUri`, `ClientID`, `Scope`, and `RedirectUri`.
     - Opens the browser (`Start-Process`) for the user to authorize.
     - Prompts you to paste the `code` from the redirect URL.
     - Exchanges that code for access/refresh tokens and saves them back into `config.json`.

3. **Zoho Header**
   - Builds `zohoHeader` with:
     - `Authorization: Zoho-oauthtoken <AccessToken>`
     - `Content-Type: application/json;charset=UTF-8`

4. **Syncro Customers**
   - Uses `SyncroAPIKey` and `SyncroSubdomain` to construct `SyncroBaseURL`.
   - Retrieves customers from Syncro.
   - Builds a list of customers that have a `ZohoCustomerId` property.
   - These become the candidates for invoice syncing.

5. **Invoice Retrieval**
   - Sets `$firstOfMonth` to the 1st of the current month.
   - Retrieves **unpaid invoices** from Syncro with a date filter `since=$firstOfMonth` and `unpaid=true`.

6. **Line Items & Product Mapping**
   - For each unpaid invoice for a mapped customer:
     - Fetches invoice line items from Syncro.
     - For each line item:
       - Looks up the product in the CSV mapping.
       - Falls back to the `DEFAULT` mapping if none is found.
       - Builds a Zoho Books line item:
         - `item_id` from `ZohoItemID`
         - `quantity` from Syncro line item quantity
         - `rate` from Syncro line item price
         - Optional `description` if `IncludeDescription` is `"Yes"`.

7. **Invoice Creation in Zoho**
   - Builds a Zoho Books invoice body including:
     - `customer_id` (the ZohoCustomerId from Syncro customer property).
     - `line_items` (from mappings).
     - `payment_terms` (e.g., 30).
   - Sends a `POST` to `https://www.zohoapis.com/books/v3/invoices?organization_id=<OrgID>`.
   - Logs success/failure and tracks counters:
     - `InvoicesProcessed`
     - `InvoicesCreated`
     - `InvoicesFailed`

8. **Optional Quick Pay in Syncro**
   - If **`-SkipQuickPay` is NOT specified**:
     - Builds a payment body for Syncro:
       - `customer_id`
       - `invoice_id`
       - `amount_cents` (from invoice total)
       - `payment_method = "Quick"`
     - Sends a `POST` to `https://<subdomain>.syncromsp.com/api/v1/payments?api_key=<APIKey>`.
     - Logs and counts:
       - `PaymentsCreated`
       - `PaymentsFailed`
   - If `-SkipQuickPay` **is** specified:
     - Skips payment creation, logging that this is “testing mode.”

9. **Final Summary**
   - Logs a summary block with counts of invoices processed/created/failed and (if applicable) payments created/failed.

---

## Usage

Open a PowerShell session in the directory containing the script and configuration files.

### Normal run (create Zoho invoices and mark paid in Syncro)

```powershell
.\REAL-SyncroToBooks.ps1
```

### Testing mode (do not create payments / do not mark as paid)

This mode still creates invoices in Zoho, but **does not** create Quick payments in Syncro:

```powershell
.\REAL-SyncroToBooks.ps1 -SkipQuickPay
```

This is recommended for initial testing to ensure:

- Product mappings are correct.
- The right customers/invoices are being processed.
- Zoho invoices appear as expected.

### Scheduling with Task Scheduler

1. Open **Task Scheduler**.
2. Create a new task:
   - Set the working directory to the folder containing the script.
   - Action: `powershell.exe`
   - Arguments (example):

     ```text
     -ExecutionPolicy Bypass -File "C:\Path\To\REAL-SyncroToBooks.ps1"
     ```

3. Choose an appropriate schedule (e.g., daily, hourly).

---

## Logging

- Log file path:  
  `SyncroToBooks-Log-YYYYMMDD.txt` in the script directory.
- Each run logs:
  - Start and end times.
  - Configuration loading and validation.
  - Token refresh / OAuth flow.
  - Number of customers, invoices, and line items processed.
  - Details of each Zoho invoice and Syncro payment creation.
  - Final summary with success/failure counts.

Use these logs to troubleshoot configuration or API issues.

---

## Security & Disclaimers

- **Protect `config.json` and logs**:
  - `config.json` contains API keys and OAuth tokens.
  - Log files may include invoice and customer IDs.
  - Restrict file permissions; do not commit secrets to public source control.
- **Test in non-production / with `-SkipQuickPay` first**:
  - To avoid duplicate or incorrect payments.
  - Verify invoices in Zoho before enabling automatic Quick Pay.
- This script is **not an official tool** from Syncro or Zoho.  
  Use at your own risk and review the code before running it in production.

---

## Customization

- **Date Range**:  
  The default is “since the first of the current month”.  
  You can modify the `$firstOfMonth` logic in the script if you want a different time window.

- **Zoho Region**:  
  If your Zoho Books organization is not in the US region, update the Zoho Books API base URL in the script (`https://www.zohoapis.com`) to the appropriate regional endpoint.

- **Additional Fields**:  
  You can extend the product mapping CSV or the JSON invoice body to include more Zoho fields (taxes, custom fields, etc.) based on your needs and Zoho Books API documentation.

---
