#Requires -Version 7.4
# Make sure to fill in all the required variables before running the script
# Also make sure the AppID used corresponds to an app with sufficient permissions, as follows:
#    Files.Read.All (delegated permission for personal Microsoft accounts)

#For details on what the script does and how to run it, check: https://michev.info/blog/post/7573/storage-report-for-personal-onedrive

[CmdletBinding()] #Make sure we can use -Verbose
Param(
[switch]$IncludeVersions, #Use the IncludeVersions switch to also include item versions statistics in the output.
[switch]$ExportToExcel #Use the ExportToExcel switch to specify whether to export the output to an Excel file.
)

function processChildren {
    Param(
    #URI for the drive
    [Parameter(Mandatory=$true)][string]$URI,
    #Drive info
    [Parameter(Mandatory=$true)]$Drive)

    if ($tokenExp -lt [datetime]::Now.AddSeconds(360)) {
        Write-Verbose "Access token is about to expire, renewing..."
        Renew-Token
    }

    $children = @()
    #fetch children, make sure to handle multiple pages
    do {
        $result = Invoke-GraphApiRequest -Uri "$URI" -Verbose:$VerbosePreference
        $URI = $result.'@odata.nextLink'

        #If we are getting multiple pages, add some delay to avoid throttling
        Start-Sleep -Milliseconds 300
        $children += $result
    } while ($URI)
    if (!$children.value) { Write-Verbose "No items found, skipping..."; return }

    $out = [System.Collections.Generic.List[object]]::new();$i=0
    Write-Verbose "Processing a total of $($children.value.count) items"
    $children = $children.value
    if (!$children) { return }

    #Process items
    foreach ($file in $children) {
        $out.Add($(processItem -Drive $Drive -file $file -Verbose:$VerbosePreference))

        if ($IncludeVersions) {
            #Anti-throttling control. We don't make any additional calls unless -IncludeVersions is specified, so only add delay here
            $i++
            if ($i % 100 -eq 0) { Start-Sleep -Milliseconds 300 }
        }
    }

    #Use the comma operator to force the output as actual list instead of array
    ,($out)
}

function processItem {
    Param(
    #Drive object
    [Parameter(Mandatory=$true)]$Drive,
    #File object
    [Parameter(Mandatory=$true)]$file)

    #Determine the item type
    if ($file.driveItem.package.Type -eq "OneNote") { $itemType = "Notebook" }
    elseif ($file.driveItem.file) { $itemType = "File" }
    elseif ($file.driveItem.folder) { $itemType = "Folder" }
    else { $itemType = "Unknown" }

    #While we can fetch versions in the initial query, you need a separate query to get version file size
    if ($IncludeVersions) { #Include version details
        if ($file.versions.count -ge 2) {
            $versions = @()
            $uri = "https://graph.microsoft.com/v1.0/drive/items/$($file.driveitem.id)/versions?`$select=size&`$top=999" #Seems you can go over 999, but just in case...

            do {#handle pagination
                $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop
                $uri = $result.'@odata.nextLink'

                $versions += $result.Value
            } while ($uri)

            if ($versions) {
                $versionSize = ($versions.size | Measure-Object -Sum).Sum
                $versionCount = $versions.count
                $versionQuota = (&{If($Drive.quota) {[math]::Round(100*$versionSize / $Drive.quota.used,2)} Else {"N/A"}})
            }
            else { #No versions found, can happen because /lists/{id}/items DOES return versions for some Folder/Notebook items, whereas the corresponding /drive call does NOT
                $versionSize = (&{If($file.driveItem.file) {$file.driveItem.size} Else {$null}}) #Only stamp this on files
                $versionCount = (&{If($file.driveItem.file) {1} Else {$null}}) #Only stamp this on files
                $versionQuota = (&{If(($null -ne $file.driveItem.size) -and $Drive.Quota.used) {[math]::Round(100*$file.driveItem.size / $Drive.quota.used,2)} Else {"N/A"}})
            }
        }
        else { #single version only, no point in querying
            $versionSize = (&{If($file.driveItem.file) {$file.driveItem.size} Else {$null}}) #Only stamp this on files
            $versionCount = (&{If($file.driveItem.file) {1} Else {$null}}) #Only stamp this on files
            $versionQuota = (&{If(($null -ne $file.driveItem.size) -and $Drive.Quota.used) {[math]::Round(100*$file.driveItem.size / $Drive.quota.used,2)} Else {"N/A"}})
        }
    }
    else {
        $versionQuota = (&{If(($null -ne $file.driveItem.size) -and $Drive.Quota.used) {[math]::Round(100*$file.driveItem.size / $Drive.quota.used,2)} Else {"N/A"}})
    }

    #Prepare the output object
    $fileinfo = [ordered]@{
        Name = $file.driveItem.name
        ItemType = $itemType
        Size = (&{If($null -ne $file.driveItem.size) {$file.driveItem.size} Else {"N/A"}})
        createdDateTime = (&{If($file.driveItem.createdDateTime) {$file.driveItem.createdDateTime} Else {"N/A"}})
        lastModifiedDateTime = (&{If($file.driveItem.lastModifiedDateTime) {$file.driveItem.lastModifiedDateTime} Else {"N/A"}})
        lastModifiedBy = (&{If($file.driveItem.lastModifiedBy) { Get-Identifier $file.driveItem.lastModifiedBy } Else {"N/A"}}) #Can be missing for some items??
        Shared = (&{If($file.driveItem.shared) {"Yes"} Else {"No"}})
        ID = $file.driveItem.Id #Hide column
        InFolder = $file.driveItem.parentReference.Id #Hide column
        ItemLink = $file.driveItem.webUrl
        ItemPath = $file.webUrl
        ItemID = "https://graph.microsoft.com/v1.0/drive/items/$($file.driveitem.id)"
        "% of Drive quota" = $versionQuota
    }
    if ($IncludeVersions -and $versionCount) { $fileinfo."VersionCount" = $versionCount }
    if ($IncludeVersions -and $versionSize) { $fileinfo."VersionSize" = $versionSize }

    #handle the output
    return [PSCustomObject]$fileinfo
}

#"Borrowed" from https://stackoverflow.com/a/42275676
function buildIndex {
    Param($array,[string]$keyName)

    $index = @{}
    foreach ($row in $array) {
        $key = $row.($keyName)
        $data = $index[$key]
        if ($data -is [Collections.ArrayList]) {
            $data.add($row) >$null
        } elseif ($data) {
            $index[$key] = [Collections.ArrayList]@($data, $row)
        } else {
            $index[$key] = $row
        }
    }
    $index
}

function Get-Identifier {
    param([Parameter(Mandatory=$true)]$Id) #Whatever Graph returns for lastModifiedBy

    #Cover additional scenarios here
    if ($Id.user) {
        if ($Id.user.email) { return $Id.user.email }
        elseif ($Id.user.displayName) { return $Id.user.displayName }
        elseif ($Id.user.id) { return $Id.user.id }
        else { return "N/A" }
    }
    else { return $Id } #catch-all
}

function Renew-Token {
    # For personal OneDrive accounts, we use MSAL with authorization code flow

    # Check if we have a cached token that's still valid
    if ($global:accessToken -and $global:tokenExp -gt [datetime]::Now.AddMinutes(5)) {
        Write-Verbose "Using cached access token, valid until $global:tokenExp"
        Set-Variable -Name authHeader -Scope Global -Value @{'Authorization'="Bearer $global:accessToken";'Content-Type'='application\json'}
        return
    }

    Write-Verbose "Acquiring new access token using MSAL authorization code flow..."

    try {
        # Create public client application, use default redirect URI for native apps if none specified
        $redirectUri ??= "http://localhost"
        $Scopes = New-Object System.Collections.Generic.List[string]
        #$Scope = "https://graph.microsoft.com/.default"
        $Scope = "Files.Read.All" # don't use .default as it comes with offline_access
        $Scopes.Add($Scope)

        $app =  [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($appID).WithRedirectUri($redirectUri).WithAuthority("https://login.microsoftonline.com/consumers/")
        $app2 = $app.Build()

        $result = $app2.AcquireTokenInteractive($Scopes).ExecuteAsync().Result

        $global:accessToken = $result.AccessToken
        Set-Variable -Name tokenExp -Scope Global -Value ($result.ExpiresOn.LocalDateTime)
        Set-Variable -Name authHeader -Scope Global -Value @{'Authorization'="Bearer $global:accessToken";'Content-Type'='application\json'}

        Write-Verbose "Access token valid until $tokenExp"
    }
    catch {
        Write-Error "Failed to obtain access token: $_"
        throw
    }
}

function Invoke-GraphApiRequest {
    param(
    [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Uri,
    [bool]$RetryOnce)

    if (!$AuthHeader) { Write-Verbose "No access token found, aborting..."; throw }

    if ($MyInvocation.BoundParameters.ContainsKey("ErrorAction")) { $ErrorActionPreference = $MyInvocation.BoundParameters["ErrorAction"] }
    else { $ErrorActionPreference = "Stop" }

    try { $result = Invoke-WebRequest -Headers $AuthHeader -Uri $uri -UseBasicParsing -Verbose:$false -ErrorAction $ErrorActionPreference -ConnectionTimeoutSeconds 300 }
    catch {
        if ($null -eq $_.Exception.Response) { throw }

        switch ($_.Exception.Response.StatusCode) {
            "TooManyRequests" { #429, throttled (Too many requests)
                if ($_.Exception.Response.Headers.'Retry-After') {
                    Write-Verbose "The request was throttled, pausing for $($_.Exception.Response.Headers.'Retry-After') seconds..."
                    Start-Sleep -Seconds $_.Exception.Response.Headers.'Retry-After'
                }
                else { Write-Verbose "The request was throttled, pausing for 10 seconds"; Start-Sleep -Seconds 10 }

                #retry the query
                $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference
            }
            "ResourceNotFound|Request_ResourceNotFound" { Write-Verbose "Resource $uri not found, skipping..."; return } #404, continue
            "BadRequest" { #400, we should terminate... but stupid Graph sometimes returns 400 instead of 404
                if ($RetryOnce) { throw } #We already retried, terminate
                Write-Verbose "Received a Bad Request reply, retry after 10 seconds just because Graph sucks..."
                Start-Sleep -Seconds 10
                $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -RetryOnce
            }
            "GatewayTimeout" {
                #Do NOT retry, the error is persistent and on the server side
                Write-Verbose "The request timed out, if this happens regularly, consider increasing the timeout or updating the query to retrieve less data per run"; throw
            }
            "ServiceUnavailable" { #Should be retriable, then again, it's Microsoft...
                if ($RetryOnce) { throw } #We already retried, terminate
                if ($_.Exception.Response.Headers.'Retry-After') {
                    Write-Verbose "The request was throttled, pausing for $($_.Exception.Response.Headers.'Retry-After') seconds..."
                    Start-Sleep -Seconds $_.Exception.Response.Headers.'Retry-After'
                }
                else {
                    Write-Verbose "The service is unavailable, pausing for 10 seconds..."
                    Start-Sleep -Seconds 10
                    $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -RetryOnce
                }
            }
            "Forbidden" { Write-Verbose "Insufficient permissions to run the Graph API call, aborting..."; throw } #403, terminate
            "InvalidAuthenticationToken" { #Access token has expired
                if ($_.ErrorDetails.Message -match "Lifetime validation failed, the token is expired|Access token has expired") { #renew token, continue
                Write-Verbose "Access token is invalid, trying to renew..."
                Renew-Token

                if (!$AuthHeader) { Write-Verbose "Failed to renew token, aborting..."; throw }
                #Token is renewed, retry the query
                $result = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference
            }}
            default { throw }
        }
    }

    if ($result) {
        if ($result.Content) { return ($result.Content | ConvertFrom-Json) }
        else { return $result }
    }
    else { throw }
}

#==========================================================================
#Main script starts here
#==========================================================================

#Variables to configure
$appID = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" #the GUID of your app with Files.Read.All delegated permission
$redirectUri = "http://localhost"

<# Load MSAL binaries
#Install via NuGet
Register-PackageSource -Provider NuGet -Name nugetRepository -Location https://www.nuget.org/api/v2
Install-Package -Name Microsoft.IdentityModel.Abstractions -Source nugetRepository
Install-Package -Name Microsoft.Identity.Client -Source nugetRepository -SkipDependencies

#Load the MSAL binaries
Add-Type -LiteralPath  "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.IdentityModel.Abstractions.8.16.0\lib\net8.0\Microsoft.IdentityModel.Abstractions.dll"
Add-Type -LiteralPath  "C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.Identity.Client.4.82.1\lib\net8.0\Microsoft.Identity.Client.dll"
#>

#Load via the ExO module
Get-Module ExchangeOnlineManagement -ListAvailable -Verbose:$false | select -First 1 | select -ExpandProperty FileList | % {
    if ($_ -match "Microsoft.Identity.Client.dll|Microsoft.IdentityModel.Abstractions.dll") {
        Add-Type -Path $_
    }
}

Renew-Token

#Get user's OneDrive information
Write-Verbose "Retrieving OneDrive drive information..."

$uri = 'https://graph.microsoft.com/v1.0/me/drive'
$Drive = Invoke-GraphApiRequest -Uri $uri -Verbose:$VerbosePreference -ErrorAction Stop

if (!$Drive) { throw "Failed to retrieve OneDrive information, aborting..." }

#Add drive-level details to the output object
$Output = @()
$driveInfo = [ordered]@{
    Name = "OneDrive Root"
    ItemType = "Drive"
    Size = $Drive.quota.used
    "% of Drive quota" = (&{If($Drive.quota.total -and ($null -ne $Drive.quota.used)) {[math]::Round(100 * ($Drive.quota.used) / ($Drive.quota.total),2)} else {"N/A"} })
    createdDateTime = (&{If($Drive.createdDateTime) {$Drive.createdDateTime} Else {"N/A"}})
    lastModifiedDateTime = (&{If($Drive.lastModifiedDateTime) {$Drive.lastModifiedDateTime} Else {"N/A"}})
    Shared = "N/A"
    ItemPath = (&{If($Drive.webUrl) {$Drive.webUrl} Else {"N/A"}})
    ItemLink = (Invoke-GraphApiRequest -Uri 'https://graph.microsoft.com/v1.0/me/drive/root' -Verbose:$VerbosePreference -ErrorAction Stop).Id # Little hack for later on
    ItemID = "https://graph.microsoft.com/v1.0/me/drive"
}
$Output += [PSCustomObject]$driveInfo

#Get all items from OneDrive
Write-Verbose "Retrieving all items from OneDrive..."
if ($IncludeVersions) {
    $uri = "https://graph.microsoft.com/v1.0/me/drive/list/items?`$expand=driveItem(`$select=id,name,webUrl,parentReference,file,folder,package,shared,size,createdDateTime,lastModifiedDateTime,lastModifiedBy),versions(`$select=id)&`$select=id,driveItem,versions,webUrl&`$top=100"
}
else {
    $uri = "https://graph.microsoft.com/v1.0/me/drive/list/items?`$expand=driveItem(`$select=id,name,webUrl,parentReference,file,folder,package,shared,size,createdDateTime,lastModifiedDateTime,lastModifiedBy)&`$select=id,driveItem,versions,webUrl&`$top=5000"
}

$pOutput = processChildren -Drive $Drive -URI $uri

if ($pOutput) {
    #Correct folder size where necessary
    if ($IncludeVersions) {#Only makes sense when we include versions
        $varIndex = buildIndex -array $pOutput -keyName "InFolder" #build index for faster lookup

        #process each folder
        $pOutput | ? {$_.ItemType -in @("Folder","Notebook")} | Sort-Object -Property {$_.ItemPath.Split("/").Count} -Descending | % {
            $Items = $varIndex[$_.ID] #Get all items with the same path as the folder
            $totalItemSize = $Items | % { if ($_.VersionSize) {$_.VersionSize} else {$_.Size} } | Measure-Object -Sum | Select-Object -ExpandProperty Sum

            if ($totalItemSize) {#Check for and correct the folder size
                if (($_.size -eq "N/A") -or ($totalItemSize -gt $_.Size)) {
                    Write-Verbose "Correcting folder size for $($_.Name)..."
                    $_.Size = $totalItemSize
                }
            }

            #Redo the '% of Drive quota' calculation
            $_."% of Drive quota" = (&{If($Drive.quota.used -and ($null -ne $_.size)) {[math]::Round(100 * ($_.Size) / ($Drive.quota.used),2)} else {"N/A"} })
        }
    }

    #Add the updated output to the main object
    $Output += $pOutput
}

#Return the output
if (!$Output -or $Output.Count -le 1) { Write-Warning "No items found in OneDrive..."; }

if ($IncludeVersions) { $Output = $Output | select Name,ItemType,Shared,Size,VersionCount,VersionSize,'% of Drive quota',createdDateTime,lastModifiedDateTime,lastModifiedBy,ItemPath,InFolder,ItemLink,ItemID }
else { $Output = $Output | select Name,ItemType,Shared,Size,'% of Drive quota',createdDateTime,lastModifiedDateTime,lastModifiedBy,ItemPath,InFolder,ItemLink,ItemID }

$global:varPersonalOneDriveItems = $Output

if ($ExportToExcel) {
    Write-Verbose "Exporting the results to an Excel file..."
    # Verify module exists
    if ($null -eq (Get-Module -Name ImportExcel -ListAvailable -Verbose:$false)) {
        Write-Warning "The ImportExcel module was not found, skipping export to Excel file..."
        return
    }

    $excel = $Output `
    ` | Export-Excel -Path "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_PersonalOneDriveStorageReport.xlsx" -WorksheetName StorageReport -FreezeTopRow -AutoFilter -BoldTopRow -NoHyperLinkConversion ItemID -AutoSize -PassThru -Verbose:$false

    #$sheet = $excel.Workbook.Worksheets["StorageReport"]

    #Add the Insights sheet
    $topFiles = $Output | ? {$_.ItemType -eq "File"} | Sort-Object -Property Size -Descending | select -First 10
    $topFilesV = $Output | ? {$_.ItemType -eq "File"} | Sort-Object -Property VersionSize -Descending | select -First 10
    $topFilesVC = $Output | ? {$_.ItemType -eq "File"} | Sort-Object -Property VersionCount -Descending | select -First 10
    $topFileTypes = $Output | ? {$_.ItemType -eq "File"} | Group-Object -Property {$_.Name.Split(".")[-1]} | Select-Object Name,Count,@{n="TotalSize";e={($_.Group | Measure-Object -Property Size -Sum).Sum}} | Sort-Object -Property TotalSize -Descending | select -First 10
    $topShared = $Output | ? {$_.Shared -eq "Yes"} | Sort-Object -Property Size -Descending | select -First 10

    if ($topFiles) { #top files by size
        $topFiles | select Name,Size,'% of Drive quota',ItemLink | Export-Excel -ExcelPackage $excel -WorksheetName "Insights" -TableName "TopFiles" -TableStyle Dark8 -StartRow 2 -AutoSize -PassThru -Verbose:$false > $null
        $sheet2 = $excel.Workbook.Worksheets["Insights"]
        Set-Format -Worksheet $sheet2 -Range A1 -Value "Top 10 Files by size" -Bold -Verbose:$false
    }
    if ($topFileTypes) { #top file types by total size
        $topFileTypes | select Name,Count,TotalSize | Export-Excel -ExcelPackage $excel -WorksheetName "Insights" -TableName "TopFileTypes" -TableStyle Dark8 -StartRow 15 -AutoSize -PassThru -Verbose:$false > $null
        Set-Format -Worksheet $sheet2 -Range A14 -Value "Top 10 File Types by total item size" -Bold -Verbose:$false
    }
    if ($topShared) { #top shared files by size
        $topShared | select Name,Size,'% of Drive quota',ItemLink | Export-Excel -ExcelPackage $excel -WorksheetName "Insights" -TableName "TopSharedFiles" -TableStyle Dark8 -StartRow 28 -AutoSize -PassThru -Verbose:$false > $null
        Set-Format -Worksheet $sheet2 -Range A27 -Value "Top 10 Shared Items by size" -Bold
    }
    if ($IncludeVersions) {
        if ($topFilesVC) { #top files by version count
            $topFilesVC | select Name,VersionCount,'% of Drive quota',ItemLink | Export-Excel -ExcelPackage $excel -WorksheetName "Insights" -TableName "TopFilesVersionCount" -TableStyle Dark8 -StartRow 41 -PassThru -Verbose:$false > $null
            Set-Format -Worksheet $sheet2 -Range A40 -Value "Top 10 Files by number of versions" -Bold
        }
        if ($topFilesV) { #top files by version size
            $topFilesV | select Name,VersionSize,'% of Drive quota',ItemLink | Export-Excel -ExcelPackage $excel -WorksheetName "Insights" -TableName "TopFilesWithVersions" -TableStyle Dark8 -StartRow 54 -PassThru -Verbose:$false > $null
            Set-Format -Worksheet $sheet2 -Range A53 -Value "Top 10 Files by size with versions included" -Bold
        }
    }

    #Save the changes
    Export-Excel -ExcelPackage $excel -WorksheetName "StorageReport" -Show -Verbose:$false
    Write-Host "Excel file exported successfully to: $($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_PersonalOneDriveStorageReport.xlsx" -ForegroundColor Green
}
else {
    $csvPath = "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_PersonalOneDriveStorageReport.csv"
    $Output | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -UseCulture
    Write-Host "Results exported to: $csvPath" -ForegroundColor Green
}

Write-Verbose "Generating HTML report..."

$htmlPath = "$($PWD)\$((Get-Date).ToString('yyyy-MM-dd_HH-mm-ss'))_PersonalOneDriveStorageReport.html"

# Calculate statistics
$usagePercent = [math]::Round(100 * $Drive.quota.used / $Drive.quota.total, 2)

# Get top level folders/items sorted by size
$topItems = $Output | Where-Object {
    #!($_.Name -eq "OneDrive Root" -or ($_.ItemType -eq "Folder" -and $_.InFolder -ne $Output[0].ItemLink)) #Only include root level items, exclude subfolders
    ($_.Name -ne "OneDrive Root")
} | Sort-Object -Property Size -Descending | Select-Object -First 200

$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Storage Metrics</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, 'Helvetica Neue', sans-serif;
            background-color: #ffffff;
            color: #201f1e;
            font-size: 14px;
        }
        .container {
            padding: 48px 32px 32px 32px;
            max-width: 1920px;
            margin: 0 auto;
        }
        .header {
            margin-bottom: 24px;
        }
        .header h1 {
            font-size: 48px;
            font-weight: 600;
            margin-bottom: 8px;
            color: #201f1e;
        }
        .storage-summary {
            display: flex;
            align-items: center;
            gap: 16px;
            margin-bottom: 32px;
            padding: 16px 0;
        }
        .storage-icon {
            width: 48px;
            height: 48px;
            background-color: #0078d4;
            border-radius: 4px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 24px;
        }
        .storage-info h2 {
            font-size: 20px;
            font-weight: 600;
            color: #201f1e;
            margin-bottom: 4px;
        }
        .storage-bar-container {
            display: flex
            align-items: center;
            gap: 12px;
        }
        .storage-text {
            font-size: 20px;
            color: #323130;
            white-space: nowrap;
        }
        .storage-bar-wrapper {
            flex: 1;
            max-width: 400px;
            height: 8px;
            background-color: #edebe9;
            border-radius: 4px;
            overflow: hidden;
        }
        .storage-bar-fill {
            height: 100%;
            background-color: #00bcf2;
            transition: width 0.3s ease;
            border-radius: 4px;
        }
        .table-container {
            background: #ffffff;
            border: 1px solid #edebe9;
        }
        .table-header {
            display: grid;
            grid-template-columns: 40px 3fr 1.5fr 1fr 1fr 1fr;
            padding: 11px 12px;
            background-color: #faf9f8;
            border-bottom: 1px solid #edebe9;
            font-size: 12px;
            font-weight: 600;
            color: #323130;
        }
        .table-header div {
            display: flex;
            align-items: center;
        }
        .table-row {
            display: grid;
            grid-template-columns: 40px 3fr 1.5fr 1fr 1fr 1fr;
            padding: 8px 12px;
            border-bottom: 1px solid #edebe9;
            align-items: center;
            transition: background-color 0.1s;
        }
        .table-row:hover {
            background-color: #f3f2f1;
        }
        .item-icon {
            width: 24px;
            height: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 16px;
        }
        .item-name {
            color: #0078d4;
            text-decoration: none;
            font-size: 14px;
            cursor: pointer;
        }
        .item-name:hover {
            text-decoration: underline;
        }
        .total-size {
            font-size: 14px;
            color: #323130;
        }
        .percent-container {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .percent-text {
            font-size: 14px;
            color: #323130;
            min-width: 50px;
        }
        .percent-bar-wrapper {
            flex: 1;
            max-width: 150px;
            height: 6px;
            background-color: #edebe9;
            border-radius: 3px;
            overflow: hidden;
        }
        .percent-bar-fill {
            height: 100%;
            background-color: #00bcf2;
            border-radius: 3px;
        }
        .quota-percent {
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .quota-bar-wrapper {
            flex: 1;
            max-width: 150px;
            height: 6px;
            background-color: #edebe9;
            border-radius: 3px;
            overflow: hidden;
        }
        .quota-bar-fill {
            height: 100%;
            background-color: #00bcf2;
            border-radius: 3px;
        }
        .date-text {
            font-size: 14px;
            color: #323130;
        }
        .folder-text {
            font-size: 14px;
            color: #323130;
            white-space: nowrap;
        }
        .pagination {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 8px;
            padding: 16px;
            background-color: #faf9f8;
            border-top: 1px solid #edebe9;
        }
        .pagination button {
            padding: 6px 12px;
            background-color: #ffffff;
            border: 1px solid #8a8886;
            color: #323130;
            cursor: pointer;
            font-size: 14px;
            border-radius: 2px;
            transition: background-color 0.1s;
        }
        .pagination button:hover:not(:disabled) {
            background-color: #f3f2f1;
        }
        .pagination button:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }
        .pagination button.active {
            background-color: #0078d4;
            color: #ffffff;
            border-color: #0078d4;
        }
        .pagination-info {
            font-size: 14px;
            color: #323130;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="storage-summary">
            <div class="storage-icon">📁</div>
            <div class="storage-info">
                <h1>My files</h1>
            </div>
            <div class="storage-info">
                <div class="storage-bar-container">
                    <span class="storage-text">$([math]::Round($Drive.quota.used / 1GB, 2)) GB used of $([math]::Round($Drive.quota.total / 1GB, 2)) GB</span>
                </div>
            <div class="storage-bar-wrapper">
                <div class="storage-bar-fill" style="width: $($usagePercent)%"></div>
            </div>
            </div>
        </div>

        <div class="table-container">
            <div class="table-header">
                <div></div>
                <div>Name</div>
                <div>In Folder</div>
                <div>Total Size</div>
                <div>% of Drive Quota</div>
                <div>Last Modified</div>
            </div>
            <div id="tableBody">
"@

foreach ($item in $topItems) {
    $icon = switch ($item.ItemType) {
        "Folder" { "📁" }
        "File" { "📄" }
        "Notebook" { "📓" }
        default { "📄" }
    }

    $size = if ($item.Size -ne "N/A" -and $item.Size -gt 0) {
        if ($item.Size -ge 1GB) { "$([math]::Round($item.Size / 1GB, 2)) GB" }
        elseif ($item.Size -ge 1MB) { "$([math]::Round($item.Size / 1MB, 2)) MB" }
        elseif ($item.Size -ge 1KB) { "$([math]::Round($item.Size / 1KB, 2)) KB" }
        else { "$($item.Size) B" }
    } else { "0 B" }

    $percentOfQuota = if ($item.'% of Drive quota' -ne "N/A") { $item.'% of Drive quota' } else { 0 }

    $lastModified = if ($item.lastModifiedDateTime -ne "N/A") {
#        ([datetime]$item.lastModifiedDateTime).ToString("M/d/yyyy г. H:mm")
        $item.lastModifiedDateTime
    } else { "" }

    $inFolder = if ($item.InFolder) {
        $parentFolder = $Output | Where-Object { $_.ItemID.Replace("https://graph.microsoft.com/v1.0/drive/items/","") -eq $item.InFolder } | Select-Object -First 1
        if ($parentFolder) { "<a href=""$($parentFolder.ItemLink)"" target=""_blank"">$($parentFolder.Name)</a>" } else { "Root" }
    } else { "Root" }

    $htmlContent += @"
                <div class="table-row" data-page="page">
                    <div class="item-icon">$icon</div>
                    <div><a href="$($item.ItemLink)" class="item-name" target="_blank">$($item.Name)</a></div>
                    <div class="folder-text">$inFolder</div>
                    <div class="total-size">$size</div>
                    <div class="quota-percent">$percentOfQuota%</span>
                        <div class="quota-bar-wrapper">
                            <div class="quota-bar-fill" style="width: $($percentOfQuota)%"></div>
                        </div>
                    </div>
                    <div class="date-text">$lastModified</div>
                </div>
"@
}

$htmlContent += @"
            </div>
            <div class="pagination">
                <button id="prevBtn" onclick="changePage(-1)">Previous</button>
                <span class="pagination-info">Page <span id="currentPage">1</span> of <span id="totalPages">1</span></span>
                <button id="nextBtn" onclick="changePage(1)">Next</button>
            </div>
        </div>
    </div>
    <script>
        const itemsPerPage = 20;
        let currentPage = 1;
        const rows = document.querySelectorAll('.table-row');
        const totalPages = Math.ceil(rows.length / itemsPerPage);

        function showPage(page) {
            rows.forEach((row, index) => {
                const startIndex = (page - 1) * itemsPerPage;
                const endIndex = startIndex + itemsPerPage;
                row.style.display = (index >= startIndex && index < endIndex) ? 'grid' : 'none';
            });

            document.getElementById('currentPage').textContent = page;
            document.getElementById('totalPages').textContent = totalPages;
            document.getElementById('prevBtn').disabled = page === 1;
            document.getElementById('nextBtn').disabled = page === totalPages;
        }

        function changePage(delta) {
            const newPage = currentPage + delta;
            if (newPage >= 1 && newPage <= totalPages) {
                currentPage = newPage;
                showPage(currentPage);
            }
        }

        showPage(1);
    </script>
</body>
</html>
"@

# Save the HTML content to a file
Set-Content -Path $htmlPath -Value $htmlContent -Encoding UTF8
Write-Host "HTML report generated successfully at: $htmlPath" -ForegroundColor Green
ii $htmlPath