#===========================================================#
#   📧 Exchange Online Mailbox Report Tool v1.0              #
#   🔧 Created: November 2024                                #
#   🔄 Last Updated: 19.11.2024                             #
#===========================================================#

<#
.SYNOPSIS
    Creates a detailed report of Exchange Online mailbox details.
.DESCRIPTION
    This PowerShell script collects comprehensive information about Exchange Online mailboxes 
    (size, type, item count, last access) and creates both a CSV export and a clear summary 
    in TXT format.
.EXAMPLE
    PS> .\ExchangeMailboxReport.ps1
    (Creates a detailed mailbox report as CSV and a summary as TXT file)
.NOTES
    License: MIT
    Version: 1.0
#>

# Function to create styled headers
function Write-StyledHeader {
    param ($Text)
    $width = 60
    $border = "═" * $width
    Write-Host "`n$border" -ForegroundColor Cyan
    Write-Host ("║ " + $Text.PadRight($width - 4) + " ║") -ForegroundColor Cyan
    Write-Host "$border`n" -ForegroundColor Cyan
}

# Function for styled section headers
function Write-SectionHeader {
    param ($Text, $Icon)
    Write-Host "`n$Icon $Text $Icon" -ForegroundColor Yellow
    Write-Host ("─" * 50) -ForegroundColor DarkGray
}

# Function to format file sizes
function Format-FileSize {
    param ([long]$Size)
    if ($Size -gt 1TB) {
        return "$([math]::Round($Size / 1TB, 2)) TB"
    }
    elseif ($Size -gt 1GB) {
        return "$([math]::Round($Size / 1GB, 2)) GB"
    }
    else {
        return "$([math]::Round($Size / 1MB, 2)) MB"
    }
}

# Function to check and install the Exchange Online module
function Initialize-ExchangeModule {
    Write-Host "🔍 Checking Exchange Online Management Module..." -ForegroundColor Yellow
    
    # Check if module is installed
    $module = Get-Module -Name ExchangeOnlineManagement -ListAvailable
    
    if (-not $module) {
        Write-Host "⚠️ Exchange Online Management Module is not installed." -ForegroundColor Yellow
        $install = Read-Host "Do you want to install the module? (Y/N)"
        
        if ($install -eq 'Y') {
            try {
                Write-Host "📥 Installing Exchange Online Management Module..." -ForegroundColor Yellow
                Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
                Write-Host "✅ Module installed successfully." -ForegroundColor Green
            }
            catch {
                Write-Host "❌ Installation error: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
        else {
            Write-Host "❌ Installation cancelled. Script cannot continue." -ForegroundColor Red
            return $false
        }
    }
    else {
        Write-Host "✅ Exchange Online Management Module is installed (Version: $($module.Version))" -ForegroundColor Green
        
        # Check for available updates
        try {
            $onlineModule = Find-Module -Name ExchangeOnlineManagement
            if ($onlineModule.Version -gt $module.Version) {
                Write-Host "🔄 An update is available (New Version: $($onlineModule.Version))" -ForegroundColor Yellow
                $update = Read-Host "Do you want to update the module? (Y/N)"
                
                if ($update -eq 'Y') {
                    Update-Module -Name ExchangeOnlineManagement -Force
                    Write-Host "✅ Module updated successfully." -ForegroundColor Green
                }
            }
        }
        catch {
            Write-Host "⚠️ Warning: Could not check for updates. Continuing with current version." -ForegroundColor Yellow
        }
    }
    
    # Import module
    Import-Module ExchangeOnlineManagement
    return $true
}

# Function to test Exchange Online connection
function Test-ExchangeConnection {
    try {
        # Try to find an active Exchange Online session
        Get-PSSession | Where-Object {
            $_.ConfigurationName -eq "Microsoft.Exchange" -and 
            $_.State -eq "Opened" -and 
            $_.Availability -eq "Available"
        } | Out-Null
        
        # Additional validation by executing an Exchange command
        Get-OrganizationConfig -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

# Function to connect to Exchange Online
function Connect-ToExchangeOnline {
    # Check if already connected
    if (Test-ExchangeConnection) {
        Write-Host "✅ Existing Exchange Online connection found." -ForegroundColor Green
        return $true
    }
    
    try {
        Write-Host "🔌 No active connection found. Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowProgress $true
        Write-Host "✅ Successfully connected to Exchange Online." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "❌ Error connecting to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to get all mailbox details
function Get-AllMailboxDetails {
    try {
        Write-Host "📊 Collecting mailbox information..." -ForegroundColor Yellow
        
        # Array for results
        $mailboxDetails = @()
        $totalSize = 0
        $userMailboxCount = 0
        
        # Get all mailboxes
        $mailboxes = Get-Mailbox -ResultSize Unlimited
        $totalMailboxes = $mailboxes.Count
        $currentMailbox = 0
        
        foreach ($mailbox in $mailboxes) {
            $currentMailbox++
            Write-Progress -Activity "📫 Processing Mailboxes" -Status "Mailbox $currentMailbox of $totalMailboxes" `
                          -PercentComplete (($currentMailbox / $totalMailboxes) * 100)
            
            # Get mailbox statistics
            $stats = Get-MailboxStatistics -Identity $mailbox.UserPrincipalName
            
            # Add size to total
            if ($stats.TotalItemSize -match "\((.*) bytes\)") {
                $totalSize += [long]$matches[1]
            }
            
            # Count user mailboxes
            if ($mailbox.RecipientTypeDetails -eq "UserMailbox") {
                $userMailboxCount++
            }
            
            # Store mailbox details in custom object
            $mailboxInfo = [PSCustomObject]@{
                DisplayName = $mailbox.DisplayName
                EmailAddress = $mailbox.PrimarySmtpAddress
                MailboxType = $mailbox.RecipientTypeDetails
                TotalItemSize = $stats.TotalItemSize
                ItemCount = $stats.ItemCount
                LastLogonTime = $stats.LastLogonTime
                Database = $mailbox.Database
                Enabled = $mailbox.IsEnabled
            }
            
            # Add to array
            $mailboxDetails += $mailboxInfo
        }
        
        Write-Progress -Activity "Processing Mailboxes" -Completed
        
        # Create summary
        $summary = [PSCustomObject]@{
            TotalMailboxes = $totalMailboxes
            UserMailboxes = $userMailboxCount
            OtherMailboxes = $totalMailboxes - $userMailboxCount
            TotalSizeBytes = $totalSize
            TotalSizeGB = [math]::Round($totalSize / 1GB, 2)
            TotalSizeTB = [math]::Round($totalSize / 1TB, 2)
            MailboxDetails = $mailboxDetails
        }
        
        # Display summary
        Write-Host "`n=== 📊 Mailbox Summary ===" -ForegroundColor Cyan
        Write-Host "📬 Total Mailboxes:     $($summary.TotalMailboxes)" -ForegroundColor Green
        Write-Host "👤 User Mailboxes:      $($summary.UserMailboxes)" -ForegroundColor Green
        Write-Host "📁 Other Mailbox Types: $($summary.OtherMailboxes)" -ForegroundColor Green
        Write-Host "💾 Total Size:          $($summary.TotalSizeGB) GB ($($summary.TotalSizeTB) TB)" -ForegroundColor Green
        Write-Host "============================`n" -ForegroundColor Cyan
        
        return $summary
    }
    catch {
        Write-Host "❌ Error retrieving mailbox details: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

# Main program
function Main {
    Clear-Host
    Write-StyledHeader "📧 Exchange Online Mailbox Report Tool"
    
    # Show welcome message
    @"
    👋 Welcome to the Exchange Online Mailbox Report Tool
    🔍 This tool will:
       • Check and verify Exchange Online connection
       • Collect detailed mailbox information
       • Generate comprehensive reports
       • Provide storage analytics
"@ | Write-Host -ForegroundColor White

    # Initialize module
    Write-SectionHeader "Module Initialization" "🔧"
    if (-not (Initialize-ExchangeModule)) {
        Write-Host "❌ Module initialization failed. Exiting..." -ForegroundColor Red
        return
    }
    Write-Host "✅ Module initialization completed" -ForegroundColor Green
    
    # Establish connection
    Write-SectionHeader "Connection Status" "🔌"
    if (Connect-ToExchangeOnline) {
        Write-Host "✅ Successfully connected to Exchange Online" -ForegroundColor Green
        
        # Get mailboxes
        Write-SectionHeader "Data Collection" "📊"
        $summary = Get-AllMailboxDetails
        
        if ($summary) {
            # Display enhanced summary
            Write-StyledHeader "📈 Mailbox Analytics Summary"
            @"
    📬 Total Mailboxes:     $($summary.TotalMailboxes)
    👤 User Mailboxes:      $($summary.UserMailboxes)
    📁 Other Mailboxes:     $($summary.OtherMailboxes)
    💾 Total Storage:       $(Format-FileSize $summary.TotalSizeBytes)
"@ | Write-Host -ForegroundColor White

            # Export reports
            Write-SectionHeader "Report Generation" "📝"
            
            $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
            $exportPath = ".\MailboxReport_$timestamp.csv"
            $summaryPath = ".\MailboxSummary_$timestamp.txt"
            
            # Export to CSV
            Write-Host "📑 Generating detailed report..." -ForegroundColor Yellow
            $summary.MailboxDetails | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
            Write-Host "✅ Report exported to: $((Get-Item $exportPath).FullName)" -ForegroundColor Green
            
            # Create enhanced summary file
            @"
╔════════════════════════════════════════════════════════╗
║            Exchange Online Mailbox Summary              ║
╠════════════════════════════════════════════════════════╣
  📅 Generated: $(Get-Date -Format 'MM/dd/yyyy HH:mm:ss')

  📊 MAILBOX STATISTICS
  ────────────────────
  📬 Total Mailboxes:     $($summary.TotalMailboxes)
  👤 User Mailboxes:      $($summary.UserMailboxes)
  📁 Other Mailboxes:     $($summary.OtherMailboxes)
  💾 Total Storage:       $(Format-FileSize $summary.TotalSizeBytes)

  💡 QUICK INSIGHTS
  ────────────────
  • User Mailbox Ratio:   $([math]::Round(($summary.UserMailboxes / $summary.TotalMailboxes) * 100, 1))%
  • Avg Size per Mailbox: $(Format-FileSize ($summary.TotalSizeBytes / $summary.TotalMailboxes))

  🔍 GENERATED BY
  ────────────────
  Exchange Online Mailbox Report Tool v1.0
╚════════════════════════════════════════════════════════╝
"@ | Out-File -FilePath $summaryPath -Encoding UTF8
            Write-Host "✅ Summary exported to: $((Get-Item $summaryPath).FullName)" -ForegroundColor Green
            
            # Connection management
            Write-SectionHeader "Session Management" "🔌"
            $disconnect = Read-Host "Do you want to disconnect from Exchange Online? (Y/N)"
            if ($disconnect -eq 'Y') {
                Write-Host "🔄 Disconnecting..." -ForegroundColor Yellow
                Disconnect-ExchangeOnline -Confirm:$false
                Write-Host "✅ Successfully disconnected from Exchange Online" -ForegroundColor Green
            }
            else {
                Write-Host "ℹ️ Connection remains active" -ForegroundColor Cyan
            }
        }
    }
    
    Write-StyledHeader "🏁 Operation Completed Successfully"
    Write-Host "Thank you for using the Exchange Online Mailbox Report Tool! 👋`n" -ForegroundColor Yellow
}

# Run script with enhanced error handling
try {
    Main
}
catch {
    Write-StyledHeader "❌ Error Occurred"
    Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please check the error and try again." -ForegroundColor Yellow
}
