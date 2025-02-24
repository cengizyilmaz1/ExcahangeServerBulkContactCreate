<#
.AUTHOR
    Cengiz YILMAZ
    Microsoft MVP 
    https://cengizyilmaz.net
    1/20/2025

.SYNOPSIS
    Creates mail-enabled contacts in Exchange from CSV data,
    setting ExternalEmailAddress to the CSV "mail" value and configuring EmailAddresses (proxy)
    and MailNickname from a sequential static address generated from user-provided prefix and domain.
    
.DESCRIPTION
    This script prompts for:
      - The target OU's Distinguished Name (where mail contacts will be created).
      - The full path to a CSV file containing contact data.
      - Whether to auto-truncate Display Names longer than 64 characters.
      - The proxy address prefix (e.g., ext.mail01).
      - The proxy address domain (e.g., contoso.com).
      
    The CSV file must include these column headers (comma-delimited):
      - "Display Name"
      - "First"
      - "Last"
      - "mail"   (this value is used only for the contact's target address)
      
    For each record, the script:
      1. Validates that required fields are present and truncates Display Name to 64 characters if necessary.
      2. Uses the CSV "mail" value as the contact's ExternalEmailAddress.
      3. Generates a sequential static address from the provided prefix and domain:
             StaticAddress = [prefix] + [two-digit sequential number] + "@" + [domain]
         For example, if prefix = ext.mail01 and domain = contoso.com:
             First contact: ext.mail01@contoso.com,
             Second contact: ext.mail02@contoso.com, etc.
      4. Sets the contact's EmailAddresses property to contain only that static address
         (marked primary with uppercase "SMTP:"), and disables address policy.
      5. Derives the MailNickname by replacing "@" with "_" in the generated static address.
      6. Creates the contact using New-MailContact (target address from CSV) and updates it using Set-MailContact.
      
    This ensures that:
      - The target address is taken from the CSV.
      - Both the proxy address (EmailAddresses) and MailNickname are solely determined from the user-provided static values.
      
.NOTES
    - Run this script in the Exchange Management Shell (EMS).
    - Ensure you have the necessary permissions to create mail contacts.
#>

# Ensure required cmdlets are available
if (-not (Get-Command New-MailContact -ErrorAction SilentlyContinue)) {
    Write-Error "New-MailContact cmdlet is not available. Run this script in the Exchange Management Shell."
    exit
}
if (-not (Get-Command Set-MailContact -ErrorAction SilentlyContinue)) {
    Write-Error "Set-MailContact cmdlet is not available. Run this script in the Exchange Management Shell."
    exit
}

# Prompt for inputs
$OU = Read-Host "Enter the target OU's Distinguished Name (e.g., OU=Contacts,DC=fixcloud,DC=com,DC=tr)"
$CSVPath = Read-Host "Enter the full path to the CSV file (e.g., C:\Data\contacts.csv)"
if (-not (Test-Path $CSVPath)) {
    Write-Error "CSV file not found: $CSVPath"
    exit
}

do {
    $autoTruncateResponse = Read-Host "Auto-truncate Display Names longer than 64 characters? (Y/N)"
    $autoTruncateResponse = $autoTruncateResponse.Trim().ToUpper()
    if ($autoTruncateResponse -ne 'Y' -and $autoTruncateResponse -ne 'N') {
        Write-Host "Invalid input. Enter Y for Yes or N for No." -ForegroundColor Yellow
    }
} until ($autoTruncateResponse -eq 'Y' -or $autoTruncateResponse -eq 'N')
$AutoTruncate = ($autoTruncateResponse -eq 'Y')
$MaxNameLength = 64

# Prompt for static address settings (for proxy and MailNickname)
$ProxyPrefix = Read-Host "Enter the proxy address prefix (e.g., ext.mail)"
$ProxyDomain = Read-Host "Enter the proxy address domain (e.g., contoso.com)"
# We'll generate sequential numbers for each contact.
$Counter = 1

# Import CSV (assuming comma-delimited)
try {
    $contacts = Import-Csv -Path $CSVPath -Delimiter "," -ErrorAction Stop
} catch {
    Write-Host "Error reading CSV file: $_" -ForegroundColor Red
    exit
}
if ($contacts.Count -eq 0) {
    Write-Host "No contacts found in the CSV file." -ForegroundColor Yellow
    exit
}

Write-Host "`nOperation starting..." -ForegroundColor Cyan

# Initialize summary counters
$TotalProcessed = 0
$SuccessfulAdds = 0
$FailedAdds = 0
$FailedResults = @()

foreach ($contact in $contacts) {
    $TotalProcessed++
    
    # Retrieve CSV fields
    $DisplayName = $contact.'Display Name'
    $FirstName   = $contact.First
    $LastName    = $contact.Last
    # The target address is taken from the CSV "mail" value.
    $CSVTarget = $contact.mail

    if ([string]::IsNullOrWhiteSpace($DisplayName) -or [string]::IsNullOrWhiteSpace($CSVTarget)) {
        Write-Host "Skipping record due to missing Display Name or mail address." -ForegroundColor Yellow
        continue
    }
    
    # Truncate Display Name if necessary
    if ($DisplayName.Length -gt $MaxNameLength) {
        if ($AutoTruncate) {
            $OriginalDisplayName = $DisplayName
            $DisplayName = $DisplayName.Substring(0, $MaxNameLength)
            Write-Host "Display Name truncated from '$OriginalDisplayName' to '$DisplayName'" -ForegroundColor Yellow
        } else {
            Write-Host "Skipping contact '$DisplayName' due to excessive length." -ForegroundColor Yellow
            continue
        }
    }
    
    # Generate sequential static address for proxy from user-provided prefix/domain
    $seqNumber = $Counter.ToString("D2")
    $StaticAddress = "$ProxyPrefix$seqNumber@$ProxyDomain"
    $Counter++
    
    # For this contact:
    #   - ExternalEmailAddress (target) is the CSV value.
    $TargetAddress = $CSVTarget
    
    # Derive MailNickname from the generated static address (replace "@" with "_")
    $MailNickname = $StaticAddress -replace "@", "_"
    # Use MailNickname as the Alias
    $Alias = $MailNickname

    try {
        # Create the mail contact with ExternalEmailAddress set to CSV target address.
        New-MailContact -Name $DisplayName `
                        -ExternalEmailAddress $TargetAddress `
                        -FirstName $FirstName `
                        -LastName $LastName `
                        -Alias $Alias `
                        -OrganizationalUnit $OU `
                        -ErrorAction Stop
                        
        # Set the contact's EmailAddresses property explicitly to contain only the generated static address
        # Mark it as primary with uppercase "SMTP:".
        $PrimaryAddress = "SMTP:" + $StaticAddress
        # Disable address policy to prevent default addresses.
        Set-MailContact -Identity $Alias -EmailAddresses @($PrimaryAddress) -EmailAddressPolicyEnabled $false -ErrorAction Stop
        
        # Update MailNickname (if supported)
        try {
            Set-MailContact -Identity $Alias -MailNickname $MailNickname -ErrorAction Stop
        } catch {
            Write-Host "Warning: Unable to update MailNickname for '$DisplayName'." -ForegroundColor Yellow
        }
        
        Write-Host "Success: Created mail contact '$DisplayName'" -ForegroundColor Green
        Write-Host "         ExternalEmailAddress (target) = $TargetAddress" -ForegroundColor Green
        Write-Host "         EmailAddresses (proxy)      = $StaticAddress" -ForegroundColor Green
        Write-Host "         MailNickname (alias)        = $MailNickname" -ForegroundColor Green
        $SuccessfulAdds++
    } catch {
        Write-Host "Error: Failed to create mail contact '$DisplayName'. Error: $($_.Exception.Message)" -ForegroundColor Red
        $FailedAdds++
        $FailedResults += [PSCustomObject]@{
            'Display Name' = $DisplayName
            'Error Message' = $($_.Exception.Message)
        }
    }
}

Write-Host ""
Write-Host "----------------------------------------" -ForegroundColor Cyan
Write-Host "Operation Summary:" -ForegroundColor Cyan
Write-Host "Total contacts processed: $TotalProcessed" -ForegroundColor Cyan
Write-Host "Successfully created: $SuccessfulAdds" -ForegroundColor Green
Write-Host "Failed to create: $FailedAdds" -ForegroundColor Red
if ($FailedAdds -gt 0) {
    Write-Host "`nFailed items:" -ForegroundColor Red
    $FailedResults | Format-Table -AutoSize
}
Write-Host "----------------------------------------" -ForegroundColor Cyan
Write-Host "Operation completed." -ForegroundColor Cyan
