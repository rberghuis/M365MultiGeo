<#
.SYNOPSIS
Retrieves the Retention Compliance Policies applied to the mailboxes

.DESCRIPTION
This script retrieves the Retention Compliance Policies applied to the mailboxes. Leverages Purview to retrieve the Retention Compliance Policy names
Connects to Exchnge Online to retrieve the mailboxes, and the hold tracking logs to determine the Retention Compliance Policies applied to the mailboxes.

Copyright (c) 2024 Robbert Berghuis | https://www.linkedin.com/in/robbertberghuis

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

The use of the third-party software links on this website is done at your own discretion and risk and with agreement that you will be solely responsible for any damage to your computer system or loss of data that results from such activities. You are solely responsible for adequate protection and backup of the data and equipment used in connection with any of the software linked to this website, and we will not be liable for any damages that you may suffer connection with downloading, installing, using, modifying or distributing such software. No advice or information, whether oral or written, obtained by you from us or from this website shall create any warranty for the software.

.INPUTS
None

.OUTPUTS
Arraylist of mailbox objects with the names of the Retention Compliance Policies applied

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
.\Get-ExoRetentioCompliancePolicy.ps1

.LINK
https://www.linkedin.com/in/robbertberghuis
https://github.com/rberghuis/M365MultiGeo
https://opensource.org/license/mit

#>

# Connect to Exchange Online and Purview
Connect-ExchangeOnline
Connect-IPPSSession -CommandName Get-RetentionCompliancePolicy, Get-AppRetentionCompliancePolicy

#region Fetch retention policies
Try {
    $OrganizationConfig = Get-OrganizationConfig -ErrorAction Stop
} Catch {
    Write-Error "Could not retrieve the Exchange Online organization configuration. Error: $($_.Exception.Message)"
    return
}
#endregion

#Fetch mailboxes
Try {
    $hitlist = Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
} Catch {
    Write-Error "Could not retrieve the Exchange Online mailboxes. Error: $($_.Exception.Message)"
    return
}

#region Retrieve retention policies
[System.Collections.ArrayList]$InPlaceHolds = @()
Foreach ($IPH in $OrganizationConfig.InPlaceHolds) { 
    $Param = [ordered]@{
        "Tenant" = $OrganizationConfig.OrganizationalUnitRoot
        "InPlaceHold" = $IPH
        "InPlaceHoldGUID" = [System.GUID]::Parse($IPH.SubString(3).Split(':')[0])
        "RetentionCompliancePolicyName" = [string]
    }
    Try {
        $Param.RetentionCompliancePolicyName = (Get-RetentionCompliancePolicy $Param.InPlaceHoldGUID -ErrorAction Stop | Select-Object -ExpandProperty Name)
    } Catch {
        $Param.RetentionCompliancePolicyName = (Get-AppRetentionCompliancePolicy $Param.InPlaceHoldGUID -ErrorAction Stop | Select-Object -ExpandProperty Name)
    }
    $InPlaceHolds.Add( (New-Object PSObject -Property $Param) ) | Out-Null
}
#endregion

#region Validate retention policies applied for each of the mailboxes
Foreach ($hit in $hitList) {
    # Defaulting variables
    $HoldTracking = @()

    # Add a custom property to the object
    $hit | Add-Member -MemberType NoteProperty -Name 'RetentionCompliancePolicy' -Value @() -Force
    
    # Fetch the hold tracking logs
    Try {
        $HoldTracking = Export-MailboxDiagnosticLogs -ComponentName HoldTracking -Identity $hit.PrimarySmtpAddress | Select-Object -ExpandProperty MailboxLog | ConvertFrom-Json
    } Catch {
        Write-Error "Failed to fetch HoldTracking logs for $($hit.PrimarySmtpAddress)" -ErrorAction Continue
    }

    # Only if there are hold tracking logs
    If ($HoldTracking.Count -ne 0) {
        # Add a custom property to the object and populate it with the retention policy name
        Foreach ($HT in $HoldTracking) {
            $HT | Add-Member -MemberType NoteProperty -Name 'RetentionCompliancePolicyName' -Value $null -Force
            $HT.RetentionCompliancePolicyName = ($InPlaceHolds | Where-Object { $_.InPlaceHold -eq $HT.hid }).RetentionCompliancePolicyName 
        }

        # Add the retention policy name to the mailbox object
        $hit.RetentionCompliancePolicy = $HoldTracking.RetentionCompliancePolicyName | Where-Object { $_ -ne $null } | Sort-Object -Unique -Descending
    }
}
#endregion

# Output for validations
$hitlist | Select-Object PrimarySmtpAddress, RetentionCompliancePolicy

# Return results
return $hitList