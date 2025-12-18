#===========================================================#
#   📧 Exchange Online Mailbox Report Tool v2.1              #
#   🔧 Created: November 2024                                #
#   🔄 Last Updated: December 2025                           #
#===========================================================#

<#
.SYNOPSIS
    Creates a detailed report of Exchange Online mailbox details.
.DESCRIPTION
    Collects comprehensive information about Exchange Online mailboxes
    and exports a CSV report plus a TXT summary.
    Uses modern browser-based authentication (default EXO V3).
.NOTES
    License: MIT
    Version: 2.1
    - Modern browser authentication only
    - Compatible with ExchangeOnlineManagement v3.x+
    - MFA supported
#>

# =========================
# UI FUNCTIONS
# =========================

function Write-StyledHeader {
    param ($Text)
    $width = 60
    $border = "═" * $width
    Write-Host "`n$border" -ForegroundColor Cyan
    Write-Host ("║ " + $Text.PadRight($width - 4) + " ║") -ForegroundColor Cyan
    Write-Host "$border`n" -ForegroundColor Cyan
}

function Write-SectionHeader {
    param ($Text, $Icon)
    Write-Host "`n$Icon $Text $Icon" -ForegroundColor Yellow
    Write-Host ("─" * 50) -ForegroundColor DarkGray
}

function Format-FileSize {
    param ([long]$Size)
    if ($Size -ge 1TB) { return "$([math]::Round($Size / 1TB, 2)) TB" }
    elseif ($Size -ge 1GB) { return "$([math]::Round($Size / 1GB, 2)) GB" }
    else { return "$([math]::Round($Size / 1MB, 2)) MB" }
}

# =========================
# MODULE INITIALIZATION
# =========================

function Initialize-ExchangeModule {
    Write-Host "🔍 Checking Exchange Online Management Module..." -ForegroundColor Yellow

    $module = Get-Module ExchangeOnlineManagement -ListAvailable |
              Sort-Object Version -Descending |
              Select-Object -First 1

    if (-not $module) {
        Write-Host "❌ ExchangeOnlineManagement module not found." -ForegroundColor Red
        Write-Host "Installing module..." -ForegroundColor Yellow
        try {
            Install-Module ExchangeOnlineManagement -Force -AllowClobber
        }
        catch {
            Write-Host "❌ Module installation failed: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "✅ Module loaded (Version: $((Get-Module ExchangeOnlineManagement).Version))" -ForegroundColor Green
    return $true
}

# =========================
# CONNECTION HANDLING
# =========================

function Test-ExchangeConnection {
    try {
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Connect-ToExchangeOnline {
    if (Test-ExchangeConnection) {
        Write-Host "✅ Existing Exchange Online connection detected." -ForegroundColor Green
        return $true
    }

    Write-Host "🔌 Connecting to Exchange Online..." -ForegroundColor Yellow
    Write-Host "🌐 Browser-based authentication will open..." -ForegroundColor Cyan

    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Start-Sleep -Seconds 2

        if (Test-ExchangeConnection) {
            $org = Get-OrganizationConfig
            Write-Host "✅ Connected to tenant: $($org.DisplayName)" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "❌ Connection verification failed." -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "❌ Connection error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# =========================
# DATA COLLECTION
# =========================

function Get-AllMailboxDetails {
    Write-Host "📊 Collecting mailbox data..." -ForegroundColor Yellow

    $mailboxes = Get-Mailbox -ResultSize Unlimited
    $results = @()
    $totalSize = 0
    $userMailboxCount = 0
    $i = 0

    foreach ($mb in $mailboxes) {
        $i++
        Write-Progress -Activity "Processing Mailboxes" `
            -Status "$i / $($mailboxes.Count)" `
            -PercentComplete (($i / $mailboxes.Count) * 100)

        $stats = Get-MailboxStatistics -Identity $mb.UserPrincipalName

        if ($stats.TotalItemSize -match '\((\d+) bytes\)') {
            $totalSize += [int64]$matches[1]
        }

        if ($mb.RecipientTypeDetails -eq 'UserMailbox') {
            $userMailboxCount++
        }

        $results += [PSCustomObject]@{
            DisplayName       = $mb.DisplayName
            EmailAddress      = $mb.PrimarySmtpAddress
            MailboxType       = $mb.RecipientTypeDetails
            TotalItemSize     = $stats.TotalItemSize
            ItemCount         = $stats.ItemCount
            LastLogonTime     = $stats.LastLogonTime
            ArchiveEnabled    = $mb.ArchiveStatus
        }
    }

    Write-Progress -Activity "Processing Mailboxes" -Completed

    return [PSCustomObject]@{
        TotalMailboxes   = $mailboxes.Count
        UserMailboxes    = $userMailboxCount
        OtherMailboxes   = $mailboxes.Count - $userMailboxCount
        TotalSizeBytes   = $totalSize
        MailboxDetails   = $results
    }
}

# =========================
# MAIN
# =========================

function Main {
    Clear-Host
    Write-StyledHeader "📧 Exchange Online Mailbox Report Tool v2.1"

    Write-SectionHeader "Module Initialization" "🔧"
    if (-not (Initialize-ExchangeModule)) { return }

    Write-SectionHeader "Connection Status" "🔌"
    if (-not (Connect-ToExchangeOnline)) { return }

    Write-SectionHeader "Data Collection" "📊"
    $summary = Get-AllMailboxDetails

    Write-StyledHeader "📈 Mailbox Summary"
    Write-Host "📬 Total Mailboxes: $($summary.TotalMailboxes)"
    Write-Host "👤 User Mailboxes:  $($summary.UserMailboxes)"
    Write-Host "📁 Other Mailboxes: $($summary.OtherMailboxes)"
    Write-Host "💾 Total Storage:   $(Format-FileSize $summary.TotalSizeBytes)"

    Write-SectionHeader "Report Generation" "📝"
    $ts = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csv = ".\MailboxReport_$ts.csv"
    $txt = ".\MailboxSummary_$ts.txt"

    $summary.MailboxDetails | Export-Csv $csv -NoTypeInformation -Encoding UTF8

@"
Exchange Online Mailbox Summary
Generated: $(Get-Date)

Total Mailboxes: $($summary.TotalMailboxes)
User Mailboxes:  $($summary.UserMailboxes)
Other Mailboxes: $($summary.OtherMailboxes)
Total Storage:   $(Format-FileSize $summary.TotalSizeBytes)
"@ | Out-File $txt -Encoding UTF8

    Write-Host "✅ CSV: $csv" -ForegroundColor Green
    Write-Host "✅ TXT: $txt" -ForegroundColor Green

    Write-SectionHeader "Session Management" "🔌"
    if ((Read-Host "Disconnect from Exchange Online? (Y/N)") -eq 'Y') {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "✅ Disconnected." -ForegroundColor Green
    }

    Write-StyledHeader "🏁 Script Execution Completed"
}

# =========================
# RUN
# =========================

try {
    Main
}
catch {
    Write-Host "❌ Fatal Error: $($_.Exception.Message)" -ForegroundColor Red
}
