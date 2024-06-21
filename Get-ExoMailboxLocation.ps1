<#
.SYNOPSIS
This script will fetch the mailbox locations for all mailboxes within the provided filter and ResultSize and for those extract the Mailbox Locations to confirm all data is stored in the expected region.

.DESCRIPTION
For each of the mailboxes found within the provided filter and ResultSize, the script will extract the Mailbox Locations and confirm if all databases are stored in the expected region.
The script will output the ExternalDirectoryObjectId, MailboxRegion, MailboxRegionLastUpdateTime, IsDone and DatabasesUsed for each mailbox.
Then return the results as a list of objects.

Copyright (c) 2024 Robbert Berghuis | https://www.linkedin.com/in/robbertberghuis

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

The use of the third-party software links on this website is done at your own discretion and risk and with agreement that you will be solely responsible for any damage to your computer system or loss of data that results from such activities. You are solely responsible for adequate protection and backup of the data and equipment used in connection with any of the software linked to this website, and we will not be liable for any damages that you may suffer connection with downloading, installing, using, modifying or distributing such software. No advice or information, whether oral or written, obtained by you from us or from this website shall create any warranty for the software.

.INPUTS
None

.OUTPUTS
Arraylist of objects within the provided filter and ResultSize

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
.\Get-ExoMailboxLocation.ps1 -Filter 'RecipientTypeDetails -eq "UserMailbox"' -ResultSize 100

.LINK
https://www.linkedin.com/in/robbertberghuis
https://github.com/rberghuis/M365MultiGeo
https://opensource.org/license/mit

#>

[CmdletBinding()]
Param (
    [ValidateNotNullOrEmpty()]
    [string]$Filter = "",

    [ValidateNotNullOrEmpty()]
    [ValidateScript({ [bool]((($_ -as [int]) -and ($_ -in 1..[int]::MaxValue)) -or ($_ -eq "Unlimited")) })]
    $ResultSize = "Unlimited"
)

#region Fetch the tenant-wide default
Try {
    $DefaultMailboxRegion = Get-OrganizationConfig -ErrorAction Stop | Select-Object -ExpandProperty DefaultMailboxRegion
} Catch {
    # Do something smart(er) here
    Write-Error "Failed to get the tenant-wide default mailbox region. Error: $($_.Exception.Message)"
    return
}
#endregion

#region Get all targets within the filter
$Properties = @("MailboxRegion", "MailboxRegionLastUpdateTime", "MailboxLocations")
If ($Filter -eq "") {
    Try {
        $hitList = Get-ExoMailbox -Properties $Properties -ResultSize $ResultSize -ErrorAction Stop
    } Catch {
        # Do something smart(er) here
        Write-Error "Failed to get the mailbox list. Error: $($_.Exception.Message)"
        return
    }
} Else {
    Try {
        $hitList = Get-ExoMailbox -Properties $Properties -ResultSize $ResultSize -Filter [scriptblock]::Create($Filter) -ErrorAction Stop
    } Catch {
        If ($_.Exception.Message -like '*invalid filter clause*') {
            # Something happened...

            Write-Warning "Could not use the provided filter '$($Filter)'. Error: $($_.Exception.Message)"
            If (($Host.UI.PromptForChoice("Get-ExoMailbox -ResultSize $($ResultSize)", "Do you want to continue without using the filter '$($filter)'?", @('&Yes', '&No'), 0)) -ne 0) {
                # User aborted the operation
                return
            }

            # User wants to continue without the filter
            Try {
                $hitList = Get-ExoMailbox -Properties $Properties -ResultSize $ResultSize -ErrorAction Stop
            } Catch {
                # Do something smart(er) here
                Write-Error "Failed to get the mailbox list without the provided filter. Error: $($_.Exception.Message)"
                return
            }
        } Else {
            # Do something smart(er) here
            Write-Error "Failed to get the mailbox list with the provided filter. Error: $($_.Exception.Message)"
            return
        }
    }
}
#endregion

#region For all mailbox objects, exfiltrate the mailbox location details and store them in a structured way
$Param = [ordered]@{
    "id" = "" # Some ID of the mailbox location object, an auto-expanded archive can have multiple AuxArchives accordingly, but this number (then) doesn't increase. Maybe 1 is a boolean for 'available' and 0 being not available but is attached/affiliated to the object?
    "MailboxGUID" = "" # GUID of the mailbox location object
    "MailboxType" = "" # Mailbox type, can be Primary, MainArchive, AuxArchive, SubstrateExtension-Teams etc.
    "ExoDAGFQDN" = "" # Assumption, this is the FQDN of the Exchange Online Database Availability Group (DAG) or the load balancer
    "ExoDatabaseGUID" = "" # Assumption, this is the GUID of the database hosted on the DAG
}

Foreach ($hit in $hitList) {
    # Add a bunch of additional properties to the object
    $hit | Add-Member -MemberType NoteProperty -Name '_MailboxLocationObjects' -Value ([System.Collections.ArrayList]@()) -Force
    $hit | Add-Member -MemberType NoteProperty -Name '_MailboxDatabases' -Value ([System.Collections.ArrayList]@()) -Force
    $hit | Add-Member -MemberType NoteProperty -Name '_AllMailboxDatabasesRegionMoved' -Value $true -Force

    # Set MailboxRegion attribute, in case it's empty (not set)
    If ($null -eq $hit.MailboxRegion) {
        # Not set, means it has defaulted to the main region - storing on the object without writing to Exchange Online for ease of use
        $hit.MailboxRegion = $DefaultMailboxRegion
    }

    # Loop through the Mailbox Locations of all mailbox objects attached to this object
    Foreach ($ml in $hit.MailboxLocations) {
        # Example input as below, following the assumption(s): ID;MailboxGUID;MailboxType;ExoDAGFQDN;ExoDatabaseGUID
        # 1;00000000-0000-0000-0000-000000000000;Primary;eurprd01.prod.outlook.com;ffffffff-ffff-ffff-ffff-ffffffffffff

        # Create object
        $obj = New-Object PSObject -Property $Param

        # Split input to a list of strings
        $mlstr = $ml.Split(';')

        # Write param values
        $obj.ID = $mlstr[0]
        $obj.MailboxGUID = $mlstr[1]
        $obj.MailboxType = $mlstr[2]
        $obj.ExoDAGFQDN = $mlstr[3]
        $obj.ExoDatabaseGUID = $mlstr[4]

        # Add obj to output
        $hit.'_MailboxLocationObjects'.Add($obj) | Out-Null

        # Capture Database region, which is the first 3 characters of the FQDN
        $DatabaseRegion = $obj.ExoDAGFQDN.Substring(0, 3).ToUpper()

        # Add result to output
        $hit.'_MailboxDatabases'.Add($DatabaseRegion) | Out-Null

        # Validate if the detected mailbox region is the same as the desired mailbox region
        If ($hit.MailboxRegion -ne $DatabaseRegion) {
            # Only need to set this to 'false' once, to indicate not all (1 or more) databases have moved
            $hit.'_AllMailboxDatabasesRegionMoved' = $false
        }
    }
}
#endregion

# Produce output
$hitList | Select-Object ExternalDirectoryObjectId, MailboxRegion, MailboxRegionLastUpdateTime, @{ Name = "IsDone"; Expression = { $_._AllMailboxDatabasesRegionMoved }}, @{ Name = "DatabasesUsed"; Expression = { ($_._MailboxLocationObjects.ExoDAGFQDN | Sort-Object -Unique) }} | Format-Table -AutoSize

# return results
return $hitList