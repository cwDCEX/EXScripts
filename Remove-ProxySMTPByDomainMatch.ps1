#TEST SCRIPT BEFORE RUNNING ON PRODUCTION ENVIRONMENT

#Update below variables to relevant CSV file path ($CSVPath) and domain to remove from proxyaddresses ($DomainToRemove)

# Import the Exchange Management Shell module
if (-not (Get-Module -Name ExchangeManagementShell -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

# Set the path to the CSV file containing the list of users
$CSVPath = "<YOURFILEPATH>"

# Set the domain to remove from the email addresses
$DomainToRemove = "<DOMAINNAME>"

# Load the list of users from the CSV file
$Users = Import-Csv -Path $CSVPath

# Loop through each user in the CSV file
foreach ($User in $Users) {
    Write-Host "Processing user $($User.AccountName)"

    # Get the mailbox for the user
    $Mailbox = Get-Mailbox -Identity $User.AccountName -ErrorAction Continue

    if ($Mailbox) {
        # Get the list of email addresses for the mailbox
        $EmailAddresses = $Mailbox.EmailAddresses | 
            Where-Object { $_.PrefixString -ne "smtp" -or $_.SmtpAddress.ToLower() -notlike "*@$DomainToRemove" } |
            ForEach-Object { [Microsoft.Exchange.Data.ProxyAddress]::Parse($_) }

        # Update the email addresses for the mailbox
        try {
            Set-Mailbox -Identity $User.AccountName -EmailAddresses $EmailAddresses -ErrorAction Stop
        }
        catch {
            if ($_.Exception.GetType().FullName -eq "Microsoft.Exchange.Data.ProxyAddressExistsException") {
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            }
            else {
                throw $_.Exception
            }
        }
    }
    else {
        Write-Host "Error: Mailbox not found for $($User.AccountName)" -ForegroundColor Red
    }
}
