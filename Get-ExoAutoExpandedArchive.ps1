<#
.SYNOPSIS
This script will return a list of mailboxes with and auto-expandED archive.

.DESCRIPTION
The script will return a list of all mailboxes including those that are inactive, with an auto-expanded archive.
The script will return the ExternalDirectoryObjectId, UserPrincipalName and a boolean value indicating if the mailbox has an auto-expanded archive.

Copyright (c) 2024 Robbert Berghuis | https://www.linkedin.com/in/robbertberghuis

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

The use of the third-party software links on this website is done at your own discretion and risk and with agreement that you will be solely responsible for any damage to your computer system or loss of data that results from such activities. You are solely responsible for adequate protection and backup of the data and equipment used in connection with any of the software linked to this website, and we will not be liable for any damages that you may suffer connection with downloading, installing, using, modifying or distributing such software. No advice or information, whether oral or written, obtained by you from us or from this website shall create any warranty for the software.

.INPUTS
None

.OUTPUTS
Arraylist of objects within the provided filter and ResultSize

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
.\Get-AutoExpandedArchive.ps1

.LINK
https://www.linkedin.com/in/robbertberghuis
https://github.com/rberghuis/M365MultiGeo
https://opensource.org/license/mit

#>

# Fetch the mailboxes
$hitList = Get-ExoMailbox -ResultSize Unlimited -IncludeInactiveMailbox -Properties MailboxLocations

# Add a property to the hitlist object to indicate if the mailbox has an AuxArchive
Foreach ($hit in $hitlist) {
    $hit | Add-Member -MemberType NoteProperty -Name '_hasAuxArchive' -Value $false -Force
    Foreach ($ml in $hit.MailboxLocations) {
        Try {
            $obj = $ml.split(';')
            If ($obj[2] -eq 'AuxArchive') {
                $hit.'_hasAuxArchive' = $true
            }
        } Catch {
            # Do something smart(er) here...
            Write-Error "Could not determine if 'AuxArchive' is present in '$($ml)' for mailbox [$($hit.ExchangeGUID)] $($hit.PrimarySmptAddress)" -ErrorAction Continue
        }
    }
}

# Produce output
$hitList | Select-Object ExternalDirectoryObjectId, UserPrincipalName, _hasAuxArchive

# return results
return $hitList