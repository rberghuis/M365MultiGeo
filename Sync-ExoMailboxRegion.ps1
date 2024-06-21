<#
.SYNOPSIS
Prompts the user to write a property on the object with value of the MailboxRegion property. This is only useful as long as the attribute selected can be used to Filter on going forward through (e.g.) Get-ExoMailbox

.DESCRIPTION
This script provides for a method of writing a property on the object with the value of the MailboxRegion property.
This closes a gap for administrators who have a need to find all mailboxes hosted within a specific region, but do not have the property set on the object to find these as the MailboxRegion property cannot be used in a filter
It provides for filtering (on filterable properties) to fetch the objects in scope, and then write any attribute on the object with the value of the MailboxRegion property.
In the examples provided, the CustomAttribute1 property is used. If the MailboxRegion property is empty, it will write the tenant-wide default value to the attribute.

Note, this is only useful as long as the attribute selected can be used to Filter on going forward through (e.g.) Get-ExoMailbox

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

.PARAMETER TargetAttribute
Specifies the attribute the update with the MailboxRegion value of the object. This is only useful as long as the attribute selected can be used to Filter on going forward through (e.g.) Get-ExoMailbox

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

# When read, understood and confirmed. Run the script as exampled below to Sync the Mailbox Region value with CustomAttribute1 for the first 100 UserMailboxes
.\Sync-ExoMailboxRegion.ps1 -Filter 'RecipientTypeDetails -eq "UserMailbox"' -TargetAttribute "CustomAttribute1" -ResultSize 100

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
    [string]$TargetAttribute = "CustomAttribute1",

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

#region Build properties list and get all targets within the filter
[System.Collections.ArrayList]$Properties = @("MailboxRegion")
$Properties.Add($TargetAttribute) | Out-Null

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

#region Update the targetAttribute with the value of the MailboxRegion, or tenant-wide default
Foreach ($hit in $hitList) {
    Try {
        # Do not overwrite existing values
        If ($null -ne $hit.$TargetAttribute) {
            Throw "$($TargetAttribute) already set for with value '$($hit.$TargetAttribute)' - not overwriting."
        }

        # Build the expression to set the target attribute, whatever that attribute may be
        $Expression = '$hit | Set-Mailbox -ErrorAction Stop -' + $TargetAttribute
        If ($null -eq $hit.MailboxRegion) {
            # If MailboxRegion is empty, it defaults to the organization's DefaultMailboxRegion
            $Expression += ' $DefaultMailboxRegion'
        } Else {
            $Expression += ' $hit.MailboxRegion'
        }
        # Results into something like: $hit | Set-Mailbox -ErrorAction Stop -CustomAttribute1 $hit.MailboxRegion

        # Capture the current Error Count, used to track if the expression caused an error
        $ErrorCnt = $Error.Count

        # Invoke the expression and capture the output
        (Invoke-Expression $Expression) 2>&1

        # Check if the expression caused an error
        If ($Error.Count -ne $ErrorCnt) {
            Throw $Error[0]
        }
    } Catch {
        # Do something smart(er) here...
        Write-Error "Failed to update '$($hit.PrimarySmtpAddress)' ($($hit.RecipientTypeDetails)) to write $($TargetAttribute) with $($hit.MailboxRegion). Error: $($_.Exception.Message)" -ErrorAction Continue
    }
}
#endregion

return