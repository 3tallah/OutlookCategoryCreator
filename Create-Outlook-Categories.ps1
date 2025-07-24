# Script Name: Create-Outlook-Categories.ps1
# Description: This script authenticates with Microsoft Graph using client credentials
# and adds multiple custom master categories to an Outlook mailbox.
# Requirements:
# - Azure AD App Registration with API permissions: MailboxSettings.ReadWrite
# - Tenant ID, Client ID, Client Secret
# - PowerShell 5.1+ or PowerShell Core

# Input parameters
$tenantId = "<YOUR_TENANT_ID>"
$clientId = "<YOUR_CLIENT_ID>"
$clientSecret = "<YOUR_CLIENT_SECRET>"
$mailboxUser = "user@example.ae"

# Get OAuth token
$body = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}
$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Body $body
$token = $tokenResponse.access_token

# Define list of categories
$categories = @(
    "ELUAE - Internal", "Action Required", "Newsletter", "Spam", "Receipt",
    "Social", "Event Invite", "Subscription Renewal", "System Notification",
    "Autoresponse"
)

# Define preset colors (rotate through preset1 - preset25)
$colorIndex = 1
foreach ($category in $categories) {
    $color = "preset$colorIndex"
    $categoryBody = @{
        displayName = $category
        color       = $color
    } | ConvertTo-Json -Depth 2

    # Create category
    try {
        $response = Invoke-RestMethod -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/users/$mailboxUser/outlook/masterCategories" `
            -Headers @{ Authorization = "Bearer $token"; "Content-Type" = "application/json" } `
            -Body $categoryBody
        Write-Host "Created category: $category"
    }
    catch {
        Write-Warning "Failed to create category '$category': $_"
    }

    $colorIndex++
    if ($colorIndex -gt 25) { $colorIndex = 1 }
}
