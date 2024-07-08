<#
.SYNOPSIS
Re-activates an inactive mailbox in Exchange Online

.DESCRIPTION
IMPORTANT please ensure you update the used CustomAttributes according to your environment and you have 'agreed' on the use of these attributes for this particular purpose.
This script allows you to re-activate a mailbox in Exchange Online, provided it is still present and does not have an auto-expanded archive (or Large Archive, AuxArchive etc.).
The script takes the following steps:
- Connects to Exchange Online and Microsoft Graph
- Create a new mailbox using the input from the old mailbox.
- Confirms through MgGraph that the directory sync has finalized between Exchange Online and Entra ID
- Blocks sign-in, prevents managed mailbox agent from processing, moves Email Addresses to CustomAttribute2 and empties the alias list - this helps prevent the mailbox from receiving new emails
- Sets the RecipientTypeDetails to match the original mailbox
- Initiates the region move and writes the new value in CustomAttribute1

Copyright (c) 2024 Robbert Berghuis | https://www.linkedin.com/in/robbertberghuis

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

The use of the third-party software links on this website is done at your own discretion and risk and with agreement that you will be solely responsible for any damage to your computer system or loss of data that results from such activities. You are solely responsible for adequate protection and backup of the data and equipment used in connection with any of the software linked to this website, and we will not be liable for any damages that you may suffer connection with downloading, installing, using, modifying or distributing such software. No advice or information, whether oral or written, obtained by you from us or from this website shall create any warranty for the software.

.INPUTS
None

.OUTPUTS
System.Array the input Csv is returned with the new ExternalDirectoryObjectId added for each mailbox that has been 're-activated'

.PARAMETER CsvInputFile
Specifies the path to the Csv-file containing the list of inactive mailboxes to re-activate
Expected elements: DisplayName, Alias, MicrosoftOnlineServicesID, ExchangeGUID, MailboxRegion, RecipientTypeDetails

.PARAMETER CsvDelimiter
Specifies the delimiter used in the Csv-file, default is ";"

.PARAMETER CsvEncoding
Specifies the encoding used in the Csv-file, default is "UTF8"

.NOTES
Copyright (c) 2024 Robbert Berghuis | https://www.linkedin.com/in/robbertberghuis

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

The use of the third-party software links on this website is done at your own discretion and risk and with agreement that you will be solely responsible for any damage to your computer system or loss of data that results from such activities. You are solely responsible for adequate protection and backup of the data and equipment used in connection with any of the software linked to this website, and we will not be liable for any damages that you may suffer connection with downloading, installing, using, modifying or distributing such software. No advice or information, whether oral or written, obtained by you from us or from this website shall create any warranty for the software.

.EXAMPLE
# Then navigate to the script's location, below example would navigate to the Downloads folder
Set-Location (Join-Path -Path $HOME -Child "Downloads")

# It might also be required to bypass the execution policy preventing the run of any unsigned / untrusted script.
# The following cmdlet can service this for the current process (only)
Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope Process

# The code is provided as-is under the MIT license as per Notes and Description.
# Always read and understand the code before executing it

# When read, understood and confirmed. Run the script as exampled below to
.\Restore-InactiveMailbox.ps1 -CsvInputFile ".\InactiveMailboxes.csv"

.LINK
https://www.linkedin.com/in/robbertberghuis
https://github.com/rberghuis/M365MultiGeo
https://opensource.org/license/mit

#>

[CmdletBinding()]
Param (
    [ValidateNotNullOrEmpty()]
    [ValidateScript({ Get-ChildItem -Path $_ -ErrorAction Stop })]
    [string]$CsvInputFile,

    [ValidateNotNullOrEmpty()]
    [string]$CsvDelimiter = ";",

    [ValidateNotNullOrEmpty()]
    [string]$CsvEncoding = "UTF8"
)

# Connect to Exhcange Online and Microsoft Graph
Connect-ExchangeOnline
Connect-MgGraph

# Stupid way of generating a 'random' password, altough not completely random etc. it is good enough for a temporary password as we block sign-in anyway
$PW =  ConvertTo-SecureString -String (-join([char[]](33..122) | Get-Random -Count 50)) -AsPlainText -Force

# Import a CSV containing a list of inactive mailboxes you want to re-activate
$hitList = Import-Csv -Path $CsvInputFile -Delimiter $CsvDelimiter -Encoding $CsvEncoding -ErrorAction Stop

Foreach ($hit in $hitList) {
    # Clear variables
    $MOSID, $NewName, $newmbx, $UserSynced = $null, $null, $null, $false

    # Add attribute to input object
    $hit | Add-Member -MemberType NoteProperty -Name 'ExternalDirectoryObjectId' -Value '' -Force

    # Define new names and IDs
    $MOSID = 'MultiGeoMove_' +  $hit.MicrosoftOnlineServicesID
    $NewName = 'MultiGeoMove_' + $hit.Alias

    # Restore mailbox to a new object
    $newmbx = New-Mailbox -InactiveMailbox $hit.ExchangeGUID -Name $NewName -MicrosoftOnlineServicesID $MOSID -MailboxRegion $hit.MailboxRegion -Password $PW

    # Store ExternalDirectoryObjectId in the input object
    $hit.ExternalDirectoryObjectId = $newmbx.ExternalDirectoryObjectId

    # Wait for Entra ID sync
    While ($false -eq $UserSynced) {
        "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Waiting for Entra ID sync..."
        Start-Sleep -Seconds 5
        $UserSynced = Try { [bool](Get-MgUser -UserId $newmbx.ExternalDirectoryObjectId -ErrorAction Stop) } Catch { $false }
    }

    # Block sign-in, prevent managed mailbox agent from processing, move Email Addresses to CA2 and empty alias list - removing all email addresses helps to prevent receiving mail on this mailbox which was inactive
    Set-Mailbox $newmbx.ExternalDirectoryObjectId -AccountDisabled $true -ElcProcessingDisabled $true -CustomAttribute2 ($newmbx.EmailAddresses -join ',') -EmailAddresses "SMTP:$($newmbx.PrimarySmtpAddress)"

    # Set RecipientTypeDetails to match (again)
    Switch ($hit.RecipientTypeDetails) {
        "SharedMailbox" { Set-Mailbox $newmbx.ExternalDirectoryObjectId -Type Shared -WarningAction SilentlyContinue; Break }
        "RoomMailbox" { Set-Mailbox $newmbx.ExternalDirectoryObjectId -Type  Room -WarningAction SilentlyContinue; Break }
        "EquipmentMailbox" { Set-Mailbox $newmbx.ExternalDirectoryObjectId -Type Equipment -WarningAction SilentlyContinue; Break }
        Default { Set-Mailbox $newmbx.ExternalDirectoryObjectId -Type Regular -WarningAction SilentlyContinue; Break }
    }

    # Initiate region move and write new value in CustomAttribute1
    Set-Mailbox $newmbx.ExternalDirectoryObjectId -MailboxRegion EUR -CustomAttribute1 "EUR"
}

# return results
return $hitList