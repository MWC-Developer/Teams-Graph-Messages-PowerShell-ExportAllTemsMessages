# PowerShell-Graph-ExportAllTeamsMessages-getAllMessages.ps1
# This file contains the Get-AllMessages function, which is called from the main script to retrieve all messages for a given chat, handling pagination and writing output.
<#
.SYNOPSIS
  Export all Teams chat messages for a user using:
    GET /users/{id}/chats/microsoft.graph.getAllMessages

  Saves output as JSON Lines (JSONL) 

  https://learn.microsoft.com/en-us/graph/api/chats-getallmessages?view=graph-rest-1.0&tabs=http
  getAllMessages - Get all messages from all chats in which a user is a participant, including one-on-one chats, group chats, and meeting chats.

.NOTES
  - Uses app-only token (client_credentials).
  - Application permissions needed:  Least privliges: Chat.Read.All	Higher privliges: Chat.ReadWrite.All
  - Don't forget admin consent for the app permissions.
  - Paginates using @odata.nextLink.
  - Writes incrementally to disk.
#>

# =========================
# CONFIG
# =========================
$TenantId     = "YOUR_TENANT_ID"           # GUID
$ClientId     = "YOUR_APP_CLIENT_ID"       # GUID
$ClientSecret = "YOUR_APP_CLIENT_SECRET"   # Secret VALUE
$UserIdOrUPN  = "user@contoso.com"         # UPN or userId (GUID)

$OutDir   = "C:\Temp\GetAllMessagesExport\$($UserIdOrUPN.Replace('@','_'))"
$OutJsonl = Join-Path $OutDir "getAllMessages.jsonl"
$OutLog   = Join-Path $OutDir "export.log"

# getAllMessages page size hint (service may cap)
$Top = 50

# Optional OData filter (leave empty string to not send)
# Example (ONLY if supported in your environment):
# $Filter = "lastModifiedDateTime gt 2026-04-01T00:00:00Z"
$Filter = ""

# =========================
# HELPERS
# =========================
function New-FolderIfMissing([string]$Path) {
  if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null }
}

function Write-Log([string]$Text) {
  $line = "[{0}] {1}" -f (Get-Date).ToString("s"), $Text
  $line | Out-File -FilePath $OutLog -Append -Encoding UTF8
}

function Get-AppOnlyToken {
  $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
  $body = @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "https://graph.microsoft.com/.default"
    grant_type    = "client_credentials"
  }

  $resp = Invoke-RestMethod -Method POST -Uri $tokenUri -Body $body -ContentType "application/x-www-form-urlencoded"
  return $resp.access_token
}

function Append-JsonlLine {
  param(
    [Parameter(Mandatory=$true)]$Object,
    [Parameter(Mandatory=$true)][string]$Path
  )
  ($Object | ConvertTo-Json -Depth 30 -Compress) | Out-File -FilePath $Path -Append -Encoding UTF8
}

# Generic GET loop that follows @odata.nextLink
function Invoke-GraphGetAllPages {
  param(
    [Parameter(Mandatory=$true)][string]$FirstUri,
    [Parameter(Mandatory=$true)][hashtable]$Headers,
    [Parameter(Mandatory=$true)][scriptblock]$OnItem
  )

  $next = $FirstUri
  $page = 0

  while ($next) {
    $page++
    Write-Log "GET page  $page  : $next"

    try {
      $resp = Invoke-RestMethod -Method GET -Uri $next -Headers $Headers
    }
    catch {
      Write-Log "ERROR GET: $next :: $($_.Exception.Message)"
      break
    }

    if ($resp.value) {
      foreach ($item in $resp.value) {
        & $OnItem $item
      }
      Write-Log "  Items in page $page : $($resp.value.Count)"
    }
    else {
      Write-Log "  No 'value' array returned on page $page."
    }

    $next = $resp.'@odata.nextLink'
  }
}

# =========================
# MAIN
# =========================
New-FolderIfMissing $OutDir
# Reset outputs
if (Test-Path $OutJsonl) { Remove-Item $OutJsonl -Force }
if (Test-Path $OutLog)   { Remove-Item $OutLog -Force }

Write-Log "Starting getAllMessages export for: $UserIdOrUPN"

$token = Get-AppOnlyToken
$headers = @{
  Authorization = "Bearer $token"
  Accept        = "application/json"
}

# Build URI:
#   /users/{id}/chats/microsoft.graph.getAllMessages?$top=50[&$filter=...]
$baseUri = "https://graph.microsoft.com/v1.0/users/$UserIdOrUPN/chats/microsoft.graph.getAllMessages?`$top=$Top"
if ($Filter -and $Filter.Trim().Length -gt 0) {
  # Keep it simple: URL-encode the filter value
  $encodedFilter = [System.Uri]::EscapeDataString($Filter)
  $baseUri = "$baseUri&`$filter=$encodedFilter"
}

Write-Log "Initial URI: $baseUri"

# Stream each message to JSONL
$counter = 0
Invoke-GraphGetAllPages -FirstUri $baseUri -Headers $headers -OnItem {
  param($msg)
  $counter++
  Append-JsonlLine -Object $msg -Path $OutJsonl

  if (($counter % 500) -eq 0) {
    Write-Log "Wrote $counter messages..."
  }
}

Write-Log "Done. Total messages written: $counter"
Write-Host "Done. Messages written: $counter"
Write-Host "Output: $OutJsonl"
Write-Host "Log:    $OutLog"
