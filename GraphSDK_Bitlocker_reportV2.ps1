#Requires -Version 3.0
#The script requires the following permissions:
#    BitLockerKey.Read.All (required)
#    Device.Read.All (optional, needed to retrieve device details)
#    User.ReadBasic.All (optional, needed to retrieve device owner's UPN)

#For details on what the script does and how to run it, check: https://www.michev.info/blog/post/7641/more-secure-version-of-the-bitlocker-recovery-keys-export-script

[CmdletBinding(SupportsShouldProcess)] #Make sure we can use -Verbose
Param(
    [switch]$IncludeDeviceInfo, #Include details about the device associated with each BitLocker key, such as device name, operating system, compliance status, etc. Requires Device.Read.All permissions.
    [switch]$IncludeDeviceOwner, #Include details about the owner of the device associated with each BitLocker key, such as the owner's UPN. Requires User.ReadBasic.All permissions. Implies IncludeDeviceInfo.
    [switch]$DeviceReport, #Generate a device-centric report, where each entry corresponds to a device and its associated BitLocker key (if any). By default, the script generates a key-centric report, where each entry corresponds to a BitLocker key. Implies IncludeDeviceInfo and IncludeDeviceOwner.
    [switch]$AllowInsecureOutput, #By default, the script exports the BitLocker keys in a secure manner (as protected XML files) to prevent accidental exposure of sensitive information. If you want to export the keys in plain text format (e.g. CSV), set this switch.
    [string]$InputFile #If you want to generate the HTML report based on a previously exported XML file, provide the path to the XML file here. The script will ignore all other parameters and generate the HTML report based on the data in the XML file.
)

#==========================================================================
#Helper functions
#==========================================================================

function DriveType {
    Param($Drive)
    switch ($Drive) {
        1 { "operatingSystemVolume" }
        2 { "fixedDataVolume" }
        3 { "removableDataVolume" }
        4 { "unknownFutureValue" }
        Default { "Unknown" }
    }
}

#Generate an HTML report
function Generate-HTMLReport {
    Param($InputData)

    if (!$InputData) { Write-Warning "No input data provided for the report, skipping HTML report generation..."; return }

    #Replace the BitLockerRecoveryKey property with its actual value
    $InputData | % { if ($_.BitLockerRecoveryKey -and ($_.BitLockerRecoveryKey -is [securestring])) { $_.BitLockerRecoveryKey = $_.BitLockerRecoveryKey | ConvertFrom-SecureString -AsPlainText }}

    #Determine the columns to include in the report, based on the properties of the input data. We exclude some "internal" properties that are not relevant for the report, adjust the list below as needed
    $cols = $InputData[0] | select * -ExcludeProperty Id,VolumeType,AdditionalProperties,CreatedDateTime,Key,BitLockerKeyId | Get-Member -MemberType NoteProperty | select -ExpandProperty Name

$HtmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #4CAF50; color: white; cursor: pointer; user-select: none; }
        th::after { content: attr(data-sort); margin-left: 5px; font-size: 0.8em; }
        th:hover { background-color: #45a049; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .key-masked { color: #666; font-style: italic; }
        .controls { margin-bottom: 15px; }
        .controls label { margin-right: 10px; }
    </style>
    <script>
        let maskKeys = true;

        function toggleKeyMask() {
            maskKeys = document.getElementById('maskCheckbox').checked;
            const keyElements = document.querySelectorAll('.bitlocker-key');
            keyElements.forEach(el => {
                if (maskKeys) {
                    if (el.dataset.key) { el.textContent = '[Key Present]'; }
                    else { el.textContent = '[No Key Present]'; }
                    el.className = 'bitlocker-key key-masked';
                } else {
                    el.textContent = el.dataset.key;
                    el.className = 'bitlocker-key';
                }
            });
        }

        let sortDirection = {};

        function sortTable(column, colName) {
            const table = document.getElementById('reportTable');
            const tbody = table.querySelector('tbody');
            const rows = Array.from(tbody.querySelectorAll('tr'));

            // Toggle sort direction
            sortDirection[colName] = sortDirection[colName] === 'asc' ? 'desc' : 'asc';
            const isAsc = sortDirection[colName] === 'asc';

            rows.sort((a, b) => {
                let aVal = a.cells[column].textContent.trim();
                let bVal = b.cells[column].textContent.trim();

                // Try to parse as numbers
                const aNum = parseFloat(aVal);
                const bNum = parseFloat(bVal);

                let comparison = 0;
                if (!isNaN(aNum) && !isNaN(bNum)) {
                    comparison = aNum - bNum;
                } else {
                    comparison = aVal.localeCompare(bVal);
                }

                return isAsc ? comparison : -comparison;
            });

            rows.forEach(row => tbody.appendChild(row));

            // Update sort indicators
            document.querySelectorAll('th').forEach(th => th.dataset.sort = '');
            event.target.dataset.sort = isAsc ? '▲' : '▼';
        }
    </script>
</head>
<body>
    <h1>BitLocker Recovery Keys Report</h1>
    <p>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <p><strong>Please delete the file once you are done with it, as it exposes the BitLocker keys in plain text. Use the XML file instead as a secure copy.</strong></p>
    <div class="controls">
        <label><input type="checkbox" id="maskCheckbox" checked onchange="toggleKeyMask()"> Mask BitLocker Keys</label>
    </div>
    <table id="reportTable" data-sort-col="-1" data-sort-order="asc">
        <thead>
            <tr>
"@

    $colIndex = 0
    foreach ($col in $cols) {
        $HtmlReport += "<th onclick='sortTable($colIndex)'>$col</th>"
        $colIndex++
    }

    $HtmlReport += "</tr></thead><tbody>"

    foreach ($item in $InputData) {
        $HtmlReport += "<tr>"
        foreach ($col in $cols) {
            $value = $item.$col
            if ($col -eq "BitLockerRecoveryKey") {
                if ($value) { $HtmlReport += "<td><span class='bitlocker-key key-masked' data-key='$value'>[Key Present]</span></td>" }
                else { $HtmlReport += "<td><span class='bitlocker-key key-masked' data-key='$value'>[No Key Present]</span></td>" }
            } else {
                $HtmlReport += "<td>$value</td>"
            }
        }
        $HtmlReport += "</tr>"
    }

    $HtmlReport += "</tbody></table></body></html>"

    $ReportPath = "$PWD\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_BitLockerKeys"
    $HtmlReport | Out-File -FilePath "$($ReportPath).html" -Encoding UTF8
    Write-Verbose "HTML report exported to $ReportPath"
    ii "$($ReportPath).html"

    #Export to CSV as well
    $InputData | Export-CSV -nti -Path "$($ReportPath).csv" -Encoding UTF8 -UseCulture
    Write-Verbose "HTML report exported to $ReportPath"
    ii "$($ReportPath).csv"
}

#==========================================================================
#Main script starts here
#==========================================================================

if ($InputFile) {
    if (-Not (Test-Path -Path $InputFile -PathType Leaf)) { Write-Error "The specified input file does not exist!"; return }
    try { $InputData = Import-Clixml -Path $InputFile -ErrorAction Stop }
    catch { Write-Error "Failed to import the input file"; return }

    Write-Verbose "Input data imported successfully from $InputFile, generating HTML report..."
    Generate-HTMLReport -InputData $InputData
    return
}

#Handle parameter dependencies, device report requires device info and so on. We set $PSBoundParameters here, as we use it to determine the required scopes later on
if ($PSBoundParameters.ContainsKey("DeviceReport") -and $PSBoundParameters["DeviceReport"]) {
    $PSBoundParameters["IncludeDeviceInfo"] = $true
    $PSBoundParameters["IncludeDeviceOwner"] = $true
}
if ($PSBoundParameters.ContainsKey("IncludeDeviceOwner") -and $PSBoundParameters["IncludeDeviceOwner"]) {
    $PSBoundParameters["IncludeDeviceInfo"] = $true
}

#Determine the required scopes, based on the parameters passed to the script
$RequiredScopes = switch ($PSBoundParameters.Keys) {
    "IncludeDeviceInfo" { if ($PSBoundParameters["IncludeDeviceInfo"]) { "Device.Read.All" } }
    "IncludeDeviceOwner" { if ($PSBoundParameters["IncludeDeviceOwner"]) { "User.ReadBasic.All" } } #Otherwise we only get the UserId
    Default { "BitLockerKey.Read.All" }
}

Write-Verbose "Connecting to Graph API..."
Import-Module Microsoft.Graph.Identity.SignIns -Verbose:$false -ErrorAction Stop
try {
    Connect-MgGraph -Scopes $RequiredScopes -verbose:$false -ErrorAction Stop -NoWelcome
}
catch { throw $_ }

#Check if we have all the required permissions
$CurrentScopes = (Get-MgContext).Scopes
if ($RequiredScopes | ? {$_ -notin $CurrentScopes }) { Write-Error "The access token does not have the required permissions, rerun the script and consent to the missing scopes!" -ErrorAction Stop }

#If requested, retrieve the device details
if ($PSBoundParameters["IncludeDeviceInfo"]) {
    Write-Verbose "Retrieving device details..."

    $Devices = @()
    if ($PSBoundParameters["IncludeDeviceOwner"]) {
        Write-Verbose "Retrieving device owner..."
        $Devices = Get-MgDevice -All -ExpandProperty registeredOwners -ErrorAction Stop -Verbose:$false
    }
    else { $Devices = Get-MgDevice -All -ErrorAction Stop -Verbose:$false }

    #Filter out devices with ID of 00000000-0000-0000-0000-000000000000
    $Devices = $Devices | ? {$Device.Id -ne "00000000-0000-0000-0000-000000000000" -or $Device.DeviceId -ne "00000000-0000-0000-0000-000000000000"}

    if ($Devices) { Write-Verbose "Retrieved $($Devices.Count) devices" }
    else { Write-Verbose "No devices found"; continue }

    #Prepare the device object to be used later on
    if ($PSBoundParameters["DeviceReport"]) {
        $Devices | Add-Member -MemberType NoteProperty -Name "BitLockerKeyId" -Value $null
        $Devices | Add-Member -MemberType NoteProperty -Name "BitLockerRecoveryKey" -Value $null
        $Devices | Add-Member -MemberType NoteProperty -Name "BitLockerDriveType" -Value $null
        $Devices | Add-Member -MemberType NoteProperty -Name "BitLockerBackedUp" -Value $null
    }
    $Devices | % { Add-Member -InputObject $_ -MemberType NoteProperty -Name "DeviceOwner" -Value (&{if ($_.registeredOwners) { $_.registeredOwners[0].AdditionalProperties.userPrincipalName } else { "N/A" }}) }
}

#Get the list of application objects within the tenant.
$Keys = @()

#Get the list of BitLocker keys
Write-Verbose "Retrieving BitLocker keys..."
$Keys = Get-MgInformationProtectionBitlockerRecoveryKey -All -ErrorAction Stop -Verbose:$false

#Cycle through the keys and retrieve the key
Write-Verbose "Retrieving BitLocker Recovery keys..."
foreach ($Key in $Keys) {
    #Skip stale/dummy devices
    if ($Key.DeviceId -eq "00000000-0000-0000-0000-000000000000") {
        Write-Warning "BitLocker key with ID $($Key.Id) has a device ID of 00000000-0000-0000-0000-000000000000, skipping..."
        continue
    }

    #Get the BitLocker key details
    try { $RecoveryKey = Get-MgInformationProtectionBitlockerRecoveryKey -BitlockerRecoveryKeyId $Key.Id -Property key -ErrorAction Stop -Verbose:$false | select -ExpandProperty Key }
    catch {
        Write-Warning "Failed to retrieve the recovery key for BitLocker key with ID $($Key.Id). This could be due to insufficient permissions or because the device is in a RMAU."
        $RecoveryKey = $null
    }
    $Key.Key = (&{if ($RecoveryKey) { $RecoveryKey } else { "N/A" }})
    $Key | Add-Member -MemberType NoteProperty -Name "BitLockerKeyId" -Value $Key.Id
    $Key | Add-Member -MemberType NoteProperty -Name "BitLockerRecoveryKey" -Value ($Key.Key | ConvertTo-SecureString -AsPlainText -Force)
    $Key | Add-Member -MemberType NoteProperty -Name "BitLockerDriveType" -Value (DriveType $Key.VolumeType)
    $Key | Add-Member -MemberType NoteProperty -Name "BitLockerBackedUp" -Value (&{if ($Key.CreatedDateTime) { Get-Date($Key.CreatedDateTime) -Format g } else { "N/A" }})

    #If requested, include the device details
    if ($PSBoundParameters["IncludeDeviceInfo"]) {
        $Device = $Devices | ? { $Key.DeviceId -eq $_.DeviceId }
        if (!$Device) {
            Write-Warning "Device with ID $($Key.DeviceId) not found!"
            $Key | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value "Device not found"
            continue
        }
        if ($Device.Id -eq "00000000-0000-0000-0000-000000000000" -or $Device.DeviceId -eq "00000000-0000-0000-0000-000000000000") {
            Write-Warning "Stale/dummy device found for key $($Key.DeviceId), skipping..."
            $Key | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value "Stale/Dummy Device"
            continue
        }

        #If building a device report, add the BitLocker key details to the device object
        if ($PSBoundParameters["DeviceReport"]) {
            $Device.BitLockerKeyId = $Key.Id
            $Device.BitLockerRecoveryKey = ($Key.Key | ConvertTo-SecureString -AsPlainText -Force)
            $Device.BitLockerDriveType = (DriveType $Key.VolumeType)
            $Device.BitLockerBackedUp = (&{if ($Key.CreatedDateTime) { Get-Date($Key.CreatedDateTime) -Format g } else { "N/A" }})
        }

        $Key | Add-Member -MemberType NoteProperty -Name "DeviceName" -Value $Device.DisplayName
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceGUID" -Value $Device.Id #key actually used by the stupid module...
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceOS" -Value $Device.OperatingSystem
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceTrustType" -Value $Device.TrustType
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceMDM" -Value $Device.managementType #can be null! ALWAYS null when using a filter...
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceCompliant" -Value $Device.isCompliant #can be null!
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceRegistered" -Value (&{if ($Device.registrationDateTime) { Get-Date($Device.registrationDateTime) -Format g } else { "N/A" }})
        $Key | Add-Member -MemberType NoteProperty -Name "DeviceLastActivity" -Value (&{if ($Device.approximateLastSignInDateTime) { Get-Date($Device.approximateLastSignInDateTime) -Format g } else { "N/A" }})

        #If requested, include the device owner
        if ($PSBoundParameters["IncludeDeviceOwner"]) {
            $Key | Add-Member -MemberType NoteProperty -Name "DeviceOwner" -Value (&{if ($Device.registeredOwners) { $Device.registeredOwners[0].AdditionalProperties.userPrincipalName } else { "N/A" }})
        }
    }

    #Simple anti-throttling control
    Start-Sleep -Milliseconds 100
}

#If no keys or devices found, exit here
if ((!$Keys -or $Keys.Count -eq 0) -and (!$Devices -or $Devices.Count -eq 0)) {
    Write-Warning "No BitLocker keys or devices found, exiting..."
    return
}

#Export the result to CLIXML (secure, recommended) and CSV (insecure, exposes the keys in plain text, use with caution!)
$ExportPath = "$PWD\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_BitLockerKeys"
if ($PSBoundParameters["DeviceReport"]) {
    #BitLocker keys are at the front, followed by the device details. Cleaned up few "internal" properties, adjust the list below as needed
    $ExportData = $Devices | select * -ExcludeProperty "AdditionalProperties","AlternativeSecurityIds","complianceExpirationDateTime","deviceMetadata","deviceVersion","memberOf","PhysicalIds","SystemLabels","transitiveMemberOf","RegisteredOwners","RegisteredUsers"

    #Export to CLIXML
    $ExportData | Export-Clixml -Path "$($ExportPath).xml" -Encoding UTF8

    #Do the insecure export if the switch is set
    if ($AllowInsecureOutput) {
        Write-Warning "AllowInsecureOutput switch is set, exporting BitLocker keys in plain text in the CSV file! Make sure to delete the file once you are done with it, as it contains sensitive information. Use the XML file instead as a secure copy."
        Generate-HTMLReport -InputData $ExportData
    }
}
else {
    $ExportData = $Keys | select * -ExcludeProperty Id,VolumeType,AdditionalProperties,CreatedDateTime,Key

    #Export to CLIXML
    $ExportData | Export-Clixml -Path "$($ExportPath).xml" -Encoding UTF8

    #Do the insecure export if the switch is set
    if ($AllowInsecureOutput) {
        Write-Warning "AllowInsecureOutput switch is set, exporting BitLocker keys in plain text in the CSV file! Make sure to delete the file once you are done with it, as it contains sensitive information. Use the XML file instead as a secure copy."
        Generate-HTMLReport -InputData $ExportData
    }
}
Write-Verbose "Output exported to $($ExportPath).xml"