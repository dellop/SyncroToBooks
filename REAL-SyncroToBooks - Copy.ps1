<#
.SYNOPSIS
    Syncs unpaid invoices from Syncro MSP to Zoho Books
.DESCRIPTION
    This script retrieves unpaid invoices from Syncro and creates corresponding invoices in Zoho Books.
    By default, it also marks invoices as "Quick Paid" in Syncro after successful creation in Zoho.
.PARAMETER SkipQuickPay
    When specified, invoices will NOT be marked as paid in Syncro. Use this for testing.
.EXAMPLE
    .\REAL-SyncroToBooks.ps1
    Runs normally with quick pay enabled
.EXAMPLE
    .\REAL-SyncroToBooks.ps1 -SkipQuickPay
    Runs without marking invoices as paid (testing mode)
#>

param(
    [switch]$SkipQuickPay
)


################## LOGGING FUNCTION ########################
# Logging function to track script execution
$LogFile = Join-Path $PSScriptRoot "SyncroToBooks-Log-$(Get-Date -Format 'yyyyMMdd').txt"

function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )

    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"

    # Write to console with color coding
    switch ($Level) {
        "INFO"    { Write-Host $LogMessage -ForegroundColor Green }
        "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $LogMessage -ForegroundColor Red }
    }

    # Append to log file
    Add-Content -Path $LogFile -Value $LogMessage
}

################## TOKEN REFRESH FUNCTION ########################
# Function to refresh Zoho access token using refresh token
function Refresh-ZohoAccessToken {
    param(
        [Parameter(Mandatory=$true)]
        [string]$RefreshToken,

        [Parameter(Mandatory=$true)]
        [string]$ClientID,

        [Parameter(Mandatory=$true)]
        [string]$ClientSecret
    )

    Write-Log "Refreshing Zoho access token using refresh token..." -Level INFO

    try {
        $Response = Invoke-RestMethod -Uri "https://accounts.zoho.com/oauth/v2/token" -Method POST -ContentType "application/x-www-form-urlencoded" -Body @{
            client_id     = $ClientID
            client_secret = $ClientSecret
            refresh_token = $RefreshToken
            grant_type    = "refresh_token"
        }

        if ($Response.access_token) {
            Write-Log "Successfully refreshed access token" -Level INFO

            # Return object with new token and expiration
            return @{
                AccessToken = $Response.access_token
                ExpiresIn = [int]$Response.expires_in
                TokenExpiration = (Get-Date).AddSeconds([int]$Response.expires_in).ToString("o")
            }
        } else {
            throw "No access token in response"
        }
    } catch {
        Write-Log "ERROR refreshing access token: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

################## SAVE TOKENS FUNCTION ########################
# Function to save tokens back to config file
function Save-ZohoTokens {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ConfigFilePath,

        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$AccessToken,

        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$RefreshToken,

        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]
        [string]$TokenExpiration
    )

    try {
        # Skip saving if we don't have at least an access token
        if (-not $AccessToken) {
            Write-Log "WARNING: Access token is empty, skipping token save" -Level WARNING
            return
        }

        # Read current config
        $Config = Get-Content -Path $ConfigFilePath -Raw | ConvertFrom-Json

        # Update token fields
        $Config.Zoho.AccessToken = $AccessToken
        $Config.Zoho.RefreshToken = if ($RefreshToken) { $RefreshToken } else { "" }
        $Config.Zoho.TokenExpiration = if ($TokenExpiration) { $TokenExpiration } else { "" }

        # Save back to file
        $Config | ConvertTo-Json | Set-Content -Path $ConfigFilePath

        Write-Log "Tokens saved to config file" -Level INFO

    } catch {
        Write-Log "ERROR saving tokens to config: $($_.Exception.Message)" -Level ERROR
        # Don't throw - tokens are in memory and script can continue
    }
}
##################################################################

Write-Log "Script started" -Level INFO
if ($SkipQuickPay) {
    Write-Log "Quick Pay is DISABLED (testing mode)" -Level WARNING
} else {
    Write-Log "Quick Pay is ENABLED - invoices will be marked as paid in Syncro" -Level INFO
}
#########################################################

################## LOAD CONFIGURATION ########################
# Load API credentials and settings from external config file
Write-Log "Loading configuration from config.json..." -Level INFO

$ConfigFile = Join-Path $PSScriptRoot "config.json"

# Check if config file exists
if (-not (Test-Path $ConfigFile)) {
    Write-Log "ERROR: Configuration file not found at: $ConfigFile" -Level ERROR
    Write-Log "Please copy config.example.json to config.json and fill in your API credentials." -Level ERROR
    throw "Configuration file not found"
}

# Load and parse config file
try {
    $Config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json

    # Validate required Zoho settings
    $RequiredZohoSettings = @("ClientID", "Secret", "RedirectUri", "AuthorizeUri", "Scope", "OrganizationID")
    foreach ($Setting in $RequiredZohoSettings) {
        if (-not $Config.Zoho.$Setting) {
            throw "Missing required Zoho setting: $Setting"
        }
    }

    # Validate required Syncro settings
    $RequiredSyncroSettings = @("APIKey", "Subdomain")
    foreach ($Setting in $RequiredSyncroSettings) {
        if (-not $Config.Syncro.$Setting) {
            throw "Missing required Syncro setting: $Setting"
        }
    }

    Write-Log "Configuration loaded successfully" -Level INFO

} catch {
    Write-Log "ERROR loading configuration: $($_.Exception.Message)" -Level ERROR
    throw
}
##############################################################

# Load Zoho settings from config
$ClientID = $Config.Zoho.ClientID
$Secret = $Config.Zoho.Secret
$RedirectUri = $Config.Zoho.RedirectUri
$AuthorizeUri = $Config.Zoho.AuthorizeUri
$Scope = $Config.Zoho.Scope
$ZohoOrgID = $Config.Zoho.OrganizationID

#region – Authorization code grant flow

################## TOKEN MANAGEMENT ########################
# Check if we have a valid refresh token to avoid manual authorization

$RefreshToken = $Config.Zoho.RefreshToken
$TokenExpiration = $Config.Zoho.TokenExpiration
$StoredAccessToken = $Config.Zoho.AccessToken

# Determine if we should use refresh token or do manual auth
$UseRefreshToken = $false

if ($RefreshToken -and $RefreshToken -ne "") {
    Write-Log "Refresh token found in config" -Level INFO

    # Check if token is still valid
    if ($TokenExpiration -and $TokenExpiration -ne "") {
        $ExpirationTime = [DateTime]::Parse($TokenExpiration)
        $TimeUntilExpiry = $ExpirationTime - (Get-Date)

        if ($TimeUntilExpiry.TotalSeconds -gt 300) {
            # Token valid for at least 5 more minutes
            Write-Log "Stored access token is still valid" -Level INFO
            $AccessToken = $StoredAccessToken
            $UseRefreshToken = $false
        } else {
            Write-Log "Access token expired or expiring soon, refreshing..." -Level INFO
            $UseRefreshToken = $true
        }
    } else {
        Write-Log "No token expiration found, refreshing tokens..." -Level INFO
        $UseRefreshToken = $true
    }
}

# If we need to refresh or don't have tokens
if ($UseRefreshToken) {
    try {
        $TokenResult = Refresh-ZohoAccessToken -RefreshToken $RefreshToken -ClientID $ClientID -ClientSecret $Secret
        $AccessToken = $TokenResult.AccessToken
        $TokenExpiration = $TokenResult.TokenExpiration

        Write-Log "Token refreshed successfully" -Level INFO

        # Save updated tokens to config
        Save-ZohoTokens -ConfigFilePath $ConfigFile -AccessToken $AccessToken -RefreshToken $RefreshToken -TokenExpiration $TokenExpiration

    } catch {
        Write-Log "Failed to refresh token, will attempt manual authorization" -Level WARNING
        $RefreshToken = ""
    }
}

# If no refresh token or refresh failed, do manual OAuth flow
if (-not $RefreshToken -or $RefreshToken -eq "") {
    Write-Log "No valid refresh token, initiating manual OAuth flow..." -Level INFO

    # Compose authorization URL
    # access_type=offline is required to get a refresh token
    $AuthUrl = "${AuthorizeUri}?client_id=$ClientID&scope=$Scope&response_type=code&redirect_uri=$RedirectUri&access_type=offline"

    Write-Log "Opening browser for authorization..." -Level INFO
    Write-Log "Requesting offline access to enable refresh tokens" -Level INFO
    Write-Host "Opening the default browser for authorization..."
    Start-Process $AuthUrl

    # Prompt user to paste the code from the redirect URL
    $AuthCode = Read-Host "Paste the authorization code you see in the browser URL after 'code='"

    # Request Access Token
    try {
        $Response = Invoke-RestMethod -Uri "https://accounts.zoho.com/oauth/v2/token" -Method POST -ContentType "application/x-www-form-urlencoded" -Body @{
            client_id     = $ClientID
            client_secret = $Secret
            redirect_uri  = $RedirectUri
            code          = $AuthCode
            grant_type    = "authorization_code"
        }

        Write-Log "OAuth response received from Zoho" -Level INFO
        Write-Log "Response contains: access_token=$($null -ne $Response.access_token), refresh_token=$($null -ne $Response.refresh_token), expires_in=$($Response.expires_in)" -Level INFO

        $AccessToken = $Response.access_token
        $RefreshToken = $Response.refresh_token

        if ($Response.expires_in) {
            $TokenExpiration = (Get-Date).AddSeconds([int]$Response.expires_in).ToString("o")
        } else {
            $TokenExpiration = ""
        }

        if ($RefreshToken) {
            Write-Log "Successfully authenticated with Zoho Books API" -Level INFO
            Write-Log "Refresh token obtained - future runs will be fully automated" -Level INFO
        } else {
            Write-Log "Successfully authenticated with Zoho Books API" -Level INFO
            Write-Log "WARNING: No refresh token in response - next run may require manual authentication" -Level WARNING
            Write-Log "This may be a Zoho API limitation or configuration issue" -Level WARNING
        }

        # Save tokens to config for future use
        Save-ZohoTokens -ConfigFilePath $ConfigFile -AccessToken $AccessToken -RefreshToken $RefreshToken -TokenExpiration $TokenExpiration

    } catch {
        Write-Log "ERROR during OAuth authorization: $($_.Exception.Message)" -Level ERROR
        throw
    }
}

##########################################################

#endregion

# Prepare Zoho API header
$zohoHeader = @{
    "Authorization" = "Zoho-oauthtoken $($AccessToken)"
    "Content-Type"  = "application/json;charset=UTF-8"
}
####################################################


### GET CURRENT INVOICE FOR SYNCRO CUSTOMER
# Load Syncro settings from config
$SyncroAPIKey = $Config.Syncro.APIKey
$SyncroSubdomain = $Config.Syncro.Subdomain
$SyncroBaseURL = "https://$SyncroSubdomain.syncromsp.com/api/v1"

$syncroAndZohoCustomers = @()
$zohoDetails = @()

#$syncroHeader = @{
#    "Authorization" = "Bearer $($SyncroAPIKey)"
#    "Content-Type" = "application/json"
#    }

################## LOAD PRODUCT MAPPINGS FROM CSV ########################
# Load product mappings from external CSV file for easier maintenance
# CSV Format: SyncroProductID, ZohoItemID, ProductName, IncludeDescription

Write-Log "Loading product mappings from CSV..." -Level INFO

$ProductMappingFile = Join-Path $PSScriptRoot "ProductMappings.csv"

# Check if CSV file exists
if (-not (Test-Path $ProductMappingFile)) {
    Write-Log "ERROR: Product mapping file not found at: $ProductMappingFile" -Level ERROR
    Write-Log "Please ensure ProductMappings.csv exists in the same directory as this script." -Level ERROR
    throw "Product mapping file not found"
}

# Load CSV and validate structure
try {
    $ProductMappings = Import-Csv -Path $ProductMappingFile

    # Validate required columns exist
    $RequiredColumns = @("SyncroProductID", "ZohoItemID", "ProductName", "IncludeDescription")
    $CsvColumns = $ProductMappings[0].PSObject.Properties.Name

    foreach ($Column in $RequiredColumns) {
        if ($Column -notin $CsvColumns) {
            throw "Missing required column: $Column"
        }
    }

    # Create hashtable for fast lookups
    $ProductLookup = @{}
    foreach ($Mapping in $ProductMappings) {
        if ($Mapping.SyncroProductID -ne "DEFAULT") {
            $ProductLookup[$Mapping.SyncroProductID] = $Mapping
        }
    }

    # Store default mapping separately
    $DefaultMapping = $ProductMappings | Where-Object { $_.SyncroProductID -eq "DEFAULT" } | Select-Object -First 1

    Write-Log "Successfully loaded $($ProductLookup.Count) product mappings" -Level INFO

} catch {
    Write-Log "ERROR loading product mappings: $($_.Exception.Message)" -Level ERROR
    throw
}
##########################################################################

    $firstOfMonth = (get-date).tostring("yyyy-MM-01")

################## RETRIEVE SYNCRO CUSTOMERS ########################
# Get all customers from Syncro that have a Zoho Customer ID property
# This creates a mapping between Syncro and Zoho customer IDs

Write-Log "Retrieving customer list from Syncro..." -Level INFO

$fullCustomerList = Invoke-RestMethod -Uri "$SyncroBaseURL/customers?api_key=$($SyncroAPIKey)" -Method Get | Select-Object -ExpandProperty customers | Select-Object -Property id,business_name,properties
foreach ($customer in $fullCustomerList) {
    if ($customer.properties.ZohoCustomerId) {
        $syncroAndZohoCustomers += @([pscustomobject]@{
            "syncroId" = $customer.id
            "zohoId" = $customer.properties.ZohoCustomerId
            })
        }
}

Write-Log "Found $($syncroAndZohoCustomers.Count) customers with Zoho Customer IDs" -Level INFO

################## RETRIEVE UNPAID INVOICES ########################
# Get all unpaid invoices from Syncro since the first of the current month

Write-Log "Retrieving unpaid invoices from Syncro since $firstOfMonth..." -Level INFO

$latestInvoices = Invoke-RestMethod -Uri "$SyncroBaseURL/invoices?api_key=$($SyncroAPIKey)&since_updated_at=$firstOfMonth&unpaid=true" -Method Get |  Select-Object -ExpandProperty invoices

Write-Log "Retrieved $($latestInvoices.Count) unpaid invoices" -Level INFO


################## PROCESS INVOICES ########################
# For each customer with both Syncro and Zoho IDs:
#   1. Get their invoice from Syncro
#   2. Retrieve line items
#   3. Map products using CSV configuration
#   4. Create invoice in Zoho Books

$InvoicesProcessed = 0
$InvoicesCreated = 0
$InvoicesFailed = 0
$PaymentsCreated = 0
$PaymentsFailed = 0

foreach ($syncroAndZohoCustomer in $syncroAndZohoCustomers) {
    $customerInvoice = $latestInvoices | Where-Object -Property customer_id -eq $($syncroAndZohoCustomer.syncroId)
    If ($customerInvoice) {
        $InvoicesProcessed++
        Write-Log "Processing invoice $($customerInvoice.id) for customer $($syncroAndZohoCustomer.syncroId)" -Level INFO

        # Get invoice line items from Syncro
        $getInvoiceDetails = Invoke-RestMethod -Uri "$SyncroBaseURL/line_items/?invoice_id=$($customerInvoice.id)&api_key=$($SyncroAPIKey)" -Method Get | Select-Object -ExpandProperty line_items

        Write-Log "Found $($getInvoiceDetails.Count) line items for invoice $($customerInvoice.id)" -Level INFO
        # Process each line item on the invoice using CSV-based product mappings
        foreach ($getInvoiceDetail in $getInvoiceDetails) {

            # Look up the product mapping in our CSV data
            $Mapping = $null

            if ($ProductLookup.ContainsKey($getInvoiceDetail.product_id)) {
                # Found a specific mapping for this product
                $Mapping = $ProductLookup[$getInvoiceDetail.product_id]
                Write-Log "Mapped product ID $($getInvoiceDetail.product_id) to $($Mapping.ProductName)" -Level INFO
            } else {
                # Use default mapping for unmapped products
                $Mapping = $DefaultMapping
                Write-Log "Product ID $($getInvoiceDetail.product_id) not found in mappings, using default: $($Mapping.ProductName)" -Level WARNING
            }

            # Build the line item for Zoho
            $LineItem = @{
                "item_id" = $Mapping.ZohoItemID
                "quantity" = $getInvoiceDetail.quantity
                "rate" = $getInvoiceDetail.price
            }

            # Add description if specified in CSV mapping
            if ($Mapping.IncludeDescription -eq "Yes") {
                $LineItem["description"] = $getInvoiceDetail.name
            }

            # Add to Zoho details array
            $zohoDetails += @([pscustomobject]$LineItem)
        }

        ################## CONSOLIDATE DUPLICATE LINE ITEMS ########################
        # Combine line items with the same item_id, description, and rate
        # This prevents multiple entries for the same work item (e.g., same ticket, multiple time entries)

        $OriginalLineItemCount = $zohoDetails.Count

        # Group items by item_id + description + rate
        $ConsolidatedDetails = @()
        $GroupedItems = $zohoDetails | Group-Object -Property {
            # Create a composite key for grouping
            $desc = if ($_.description) { $_.description } else { "" }
            "$($_.item_id)|$desc|$($_.rate)"
        }

        foreach ($Group in $GroupedItems) {
            # Sum quantities for all items in this group
            $TotalQuantity = ($Group.Group | Measure-Object -Property quantity -Sum).Sum

            # Use the first item as a template and update the quantity
            $ConsolidatedItem = $Group.Group[0].PSObject.Copy()
            $ConsolidatedItem.quantity = $TotalQuantity

            $ConsolidatedDetails += $ConsolidatedItem
        }

        $zohoDetails = $ConsolidatedDetails

        if ($OriginalLineItemCount -ne $zohoDetails.Count) {
            Write-Log "Consolidated $OriginalLineItemCount line items into $($zohoDetails.Count) for invoice $($customerInvoice.id)" -Level INFO
        }
        ##########################################################################

        # Build the invoice body for Zoho Books API
        $body =@{
            "customer_id" = $syncroAndZohoCustomer.zohoId
			"payment_terms" = 30
            "line_items" = $zohoDetails
                }

        $bodyJson = $body | ConvertTo-Json

        Write-Log "Creating invoice in Zoho Books for customer $($syncroAndZohoCustomer.zohoId)..." -Level INFO

        # Attempt to create invoice in Zoho Books with error handling
        try {
            $ZohoResponse = Invoke-RestMethod -Uri "https://www.zohoapis.com/books/v3/invoices/?organization_id=$ZohoOrgID" -Method POST -Headers $zohoHeader -Body $bodyJson

            if ($ZohoResponse) {
                $InvoicesCreated++
                Write-Log "Successfully created invoice in Zoho Books. Zoho Invoice ID: $($ZohoResponse.invoice.invoice_id)" -Level INFO

                ################MARK INVOICE AS PAID IN SYNCRO########################
                # Create a "Quick" payment in Syncro to mark the invoice as paid
                if (-not $SkipQuickPay) {
                    Write-Log "Creating Quick payment in Syncro for invoice $($customerInvoice.id)..." -Level INFO

                    try {
                        # Build payment body according to Syncro API specification
                        $paymentBody = @{
                            "customer_id" = $customerInvoice.customer_id
                            "invoice_id" = $customerInvoice.id
                            "amount_cents" = $customerInvoice.total
                            "payment_method" = "Quick"
                        }

                        $paymentBodyJson = $paymentBody | ConvertTo-Json

                        # Create payment header
                        $syncroPaymentHeader = @{
                            "Content-Type" = "application/json"
                        }

                        # Create payment in Syncro
                        $PaymentResponse = Invoke-RestMethod -Uri "$SyncroBaseURL/payments?api_key=$($SyncroAPIKey)" -Method POST -Headers $syncroPaymentHeader -Body $paymentBodyJson

                        if ($PaymentResponse -and $PaymentResponse.payment) {
                            $PaymentsCreated++
                            Write-Log "Successfully created Quick payment in Syncro. Payment ID: $($PaymentResponse.payment.id)" -Level INFO
                        }
                    } catch {
                        $PaymentsFailed++
                        Write-Log "ERROR creating payment in Syncro: $($_.Exception.Message)" -Level ERROR
                        Write-Log "Invoice will remain unpaid in Syncro. Invoice ID: $($customerInvoice.id)" -Level WARNING
                    }
                } else {
                    Write-Log "Skipping Quick payment (testing mode)" -Level WARNING
                }
                ##################################################################
            }
        } catch {
            $InvoicesFailed++
            Write-Log "ERROR creating invoice in Zoho: $($_.Exception.Message)" -Level ERROR
            Write-Log "Invoice ID: $($customerInvoice.id), Customer: $($syncroAndZohoCustomer.syncroId)" -Level ERROR
        }

        # Clear the array for next customer
        $zohoDetails = @()
    }
}

################## FINAL SUMMARY ########################
Write-Log "======================================" -Level INFO
Write-Log "Script execution completed" -Level INFO
Write-Log "Invoices processed: $InvoicesProcessed" -Level INFO
Write-Log "Invoices created in Zoho: $InvoicesCreated" -Level INFO
Write-Log "Invoices failed: $InvoicesFailed" -Level INFO
if (-not $SkipQuickPay) {
    Write-Log "Payments created in Syncro: $PaymentsCreated" -Level INFO
    Write-Log "Payments failed: $PaymentsFailed" -Level INFO
} else {
    Write-Log "Quick Pay was disabled (testing mode)" -Level WARNING
}
Write-Log "======================================" -Level INFO