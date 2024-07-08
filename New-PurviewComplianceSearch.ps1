<#
.SYNOPSIS
Provides the code to export all items from an auto-expanded archive (only) of a specific mailbox

.DESCRIPTION
An auto-expanded archive cannot be restored using normal methods like 'New-Mailbox -InactiveMailbox' or 'New-MailboxRestoreRequest -SourceIsArchive'.
This script provides the code to export all items from an auto-expanded archive (only) of a specific mailbox.
It will retrieve the folders of the Online Archive (only) and based on that, create a KQL query to be used in a Compliance Search.
It will create a new Compliance Search, start the search, wait until the search is completed, and output the results for the admin to validate the numbers.
The script will also output the total source size based on FolderSize and the total items count of the source based on ItemsInFolder.

Copyright (c) 2024 Robbert Berghuis | https://www.linkedin.com/in/robbertberghuis

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

The use of the third-party software links on this website is done at your own discretion and risk and with agreement that you will be solely responsible for any damage to your computer system or loss of data that results from such activities. You are solely responsible for adequate protection and backup of the data and equipment used in connection with any of the software linked to this website, and we will not be liable for any damages that you may suffer connection with downloading, installing, using, modifying or distributing such software. No advice or information, whether oral or written, obtained by you from us or from this website shall create any warranty for the software.

.INPUTS
None

.OUTPUTS
None

.PARAMETER Filter
Specifies the filter to use when fetching the mailboxes. The filter should be a valid filter clause for Get-ExoMailbox. If the filter is not valid, the script will prompt the user to continue without the filter. If the user chooses to continue without the filter, the script will fetch all mailboxes. If the user chooses not to continue without the filter, the script will exit.

.PARAMETER ResultSize
Specifies the maximum number of results to return. The default value is "Unlimited". When not 'Unlimited', should be a positive integer.

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

# When read, understood and confirmed. Run the script as exampled below to get the Mailbox Locations for the first 100 UserMailboxes
.\New-PurviewComplianceSearch.ps1 -Target 'some-mailbox@contoso.com'

.LINK
https://techcommunity.microsoft.com/t5/exchange-team-blog/content-search-for-targeted-collection-of-inactive-mailbox-data/ba-p/3719422
https://learn.microsoft.com/en-us/purview/ediscovery-use-content-search-for-targeted-collections?view=o365-worldwide

https://www.linkedin.com/in/robbertberghuis
https://github.com/rberghuis/M365MultiGeo
https://opensource.org/license/mit

#>
[CmdletBinding()]
Param (
    [ValidateNotNullOrEmpty()]
    [string]$Target
)

# Fetch the mailbox targeted
$hit = Get-Mailbox -InactiveMailboxOnly $target

# Fetch the Mailbox Folder Statistics to prep an Content Search Export based on FolderIDs
# Main mailbox can be restored through Exchange Online New-MailboxRestoreRequest
# $FolderIds = Get-EXOMailboxFolderStatistics -Identity $hit.ExchangeGuid -IncludeSoftDeletedRecipients
$ArchiveFolderIds = Get-EXOMailboxFolderStatistics -Identity $hit.ExchangeGuid -IncludeSoftDeletedRecipients -Archive

# Create list for Mailbox Location objects
$objMailboxLocations = @()

# Create param list for custom object
$Param = [ordered]@{
    "id" = "" # Some ID of the mailbox location object, an auto-expanded archive can have multiple AuxArchives accordingly, but this number (then) doesn't increase. Maybe 1 is a boolean for 'available' and 0 being not available but is attached/affiliated to the object?
    "MailboxGUID" = "" # GUID of the mailbox location object
    "MailboxType" = "" # Mailbox type, can be Primary, MainArchive, AuxArchive, SubstrateExtension-Teams etc.
    "ExoDAGFQDN" = "" # Assumption, this is the FQDN of the Exchange Online Database Availability Group (DAG) or the load balancer
    "ExoDatabaseGUID" = "" # Assumption, this is the GUID of the database hosted on the DAG
    "CombinedMailboxTypeID" = "" # Combination of MailboxType and ID
}

# Foreach of the MailboxLocations
Foreach ($ml in $hit.MailboxLocations) {
    # Example input: 1;1e46b4a6-ee53-4e47-8c18-68cda44323d9;AuxArchive;apcprd04.prod.outlook.com;ac72736c-f194-4279-9734-5d4887bb2f2e
    
    # Create object
    $obj = New-Object PSObject -Property $Param

    # Write param values
    $obj.ID = $ml.Split(';')[0]
    $obj.MailboxGUID = $ml.Split(';')[1]
    $obj.MailboxType = $ml.Split(';')[2]
    $obj.ExoDAGFQDN = $ml.Split(';')[3]
    $obj.ExoDatabaseGUID = $ml.Split(';')[4]
    $obj.CombinedMailboxTypeID = "$($obj.MailboxType)-$($obj.ID)"

    # Add obj to list
    $objMailboxLocations += $obj 
}

# region  Create a folder Query to be used in Content Search
# SOURCE https://learn.microsoft.com/en-us/purview/ediscovery-use-content-search-for-targeted-collections?view=o365-worldwide

# Create list for folder Queries
$QueryList = @()

# Create param list for custom object
$Param = [ordered]@{
    "ExchangeGUID" = ""
    "ContentMailboxGuid" = ""
    "FolderPath" = ""
    "FolderQuery" = ""
    "_ArchiveType" = "" # Assumption
    "_ArchiveID" = "" # Assumption
}

# Foreach of the entries in the Folder Stats export
Foreach ($fs in $ArchiveFolderIds) {
    # Defaulting variables
    $folderId, $folderPath, $encoding, $nibbler, $folderIdBytes, $indexIdBytes, $indexIdIdx, $folderQuery = $null
    $obj = New-Object PSObject -Property $Param

    $folderId = $fs.FolderId
    $folderPath = $fs.FolderPath
    $encoding = [System.Text.Encoding]::GetEncoding("us-ascii")
    $nibbler = $encoding.GetBytes("0123456789ABCDEF")
    $folderIdBytes = [Convert]::FromBase64String($folderId)
    $indexIdBytes = New-Object byte[] 48
    $indexIdIdx = 0
    $folderIdBytes | Select-Object -Skip 23 -First 24 | Foreach-Object { $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -shr 4]; $indexIdBytes[$indexIdIdx++] = $nibbler[$_ -band 0xF] }
    $folderQuery = "folderid:$($encoding.GetString($indexIdBytes))"
    
    # Store output in object
    $obj.ExchangeGUID = $fs.Identity.Split('\')[0]
    $obj.ContentMailboxGuid = $fs.ContentMailboxGuid
    $obj.FolderPath = $folderPath
    $obj.FolderQuery = $folderQuery
    $obj.'_ArchiveType' = $objMailboxLocations | where-Object { $_.MailboxGUID -eq $fs.ContentMailboxGuid.ToString() } | Select-Object -ExpandProperty MailboxType    
    $obj.'_ArchiveID' = $objMailboxLocations | where-Object { $_.MailboxGUID -eq $fs.ContentMailboxGuid.ToString() } | Select-Object -ExpandProperty CombinedMailboxTypeID
    
    # Add obj to list
    $QueryList += $obj
}

$KQL = $QueryList.FolderQuery -Join ' OR '

# Confirm Folder ID Count equals Archive Folder count
($ArchiveFolderIds.Count -eq $QueryList.FolderQuery.Count)
"Source location count: $($ArchiveFolderIds.Count)"
"Query count: $($QueryList.FolderQuery.Count)"

# Create a new Compliance Search
$NCS = New-ComplianceSearch -Name "ArchiveExport-$($hit.Alias)" -ExchangeLocation $hit.PrimarySmtpAddress -AllowNotFoundExchangeLocationsEnabled $true -ContentMatchQuery $KQL -Description "Export of the Archive of [$($hit.ExchangeGUD)] $($hit.Alias)"

# Start the Compliance search
Start-ComplianceSearch $NCS.Name

# Wait until completed
While ($false -eq ((Get-ComplianceSearch $NCS.Name).Status -eq 'Completed')) {
    "[$(Get-Date -Format 'yyyyMMdd hh:mm:ss')] search not yet completed..."
    Start-Sleep -Seconds 300
}

# Calculate total source size based on FolderSize and item count based on ItemsInFolder
$SourceSize = (($ArchiveFolderIds | Foreach-Object { $_.FolderSize.Split('(')[1].Split(' ')[0].Replace(',', '') }) | Measure-Object -Sum).Sum
$SourceCount = ($ArchiveFolderIds.ItemsInFolder | Measure-Object -Sum).Sum
"Source location size: $($SourceSize) (bytes)"
"Source location total item count: $($SourceCount)"

# Output the Compliance Search results, to validate the numbers
$SearchResults = Get-ComplianceSearch $NCS.Name
$SearchItemCount = $SearchResults.Items + $SearchResults.UnindexedItems
$SearchSize = $SearchResults.Size + $SearchResults.UnindexedSize
"Search result size: $($SearchSize) (bytes)"
"Search result total item count: $($SearchItemCount)"
If ( ($SourceSize -ne $SearchSize) -or $($SourceCount -ne $SearchItemCount) ) {
    Write-Error "The Compliance Search results do not match the expected results. Please validate the search and the results." -ErrorAction Continue
}

return