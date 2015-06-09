<#
    .SYNOPSIS
    Generates a report for Exchange 2010 Public Folder Replication.
    
    This is an updated version of the Mike Walker (blog.mikewalker.me) to support non-ASCII environments.

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
    
    .LINK
    Original Version of the script by Mike Walker: https://gallery.technet.microsoft.com/office/Exchange-2010-Public-944df6ee

    .DESCRIPTION
    This script will generate a report for Exchange 2010 Public Folder Replication. It returns general information, such as total number of public folders, total items in all public folders, total size of all items, the top 10 largest folders, and more. Additionally, it lists each Public Folder and the replication status on each server. By default, this script will scan the entire Exchange environment in the current domain and all public folders. This can be limited by using the -ComputerName and -FolderPath parameters.
    
    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1  
    - Exchange Server 2010

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release of the updated version 

    .PARAMETER ComputerName
    This parameter specifies the Exchange 2010 server(s) to scan. If this is omitted, all Exchange servers with the Mailbox role in the current domain are scanned.

    .PARAMETER FolderPath
    This parameter specifies the Public Folder(s) to scan. If this is omitted, all public folders are scanned.

    .PARAMETER Recurse
    When used in conjunction with the FolderPath parameter, this will include all child Public Folders of the Folders listed in Folder Path.

    .PARAMETER AsHTML
    Specifying this switch will have this script output HTML, rather than the result objects. This is independent of the Filename or SendEmail parameters and only controls the console output of the script.

    .PARAMETER Filename
    Providing a Filename will save the HTML report to a file.

    .PARAMETER SendEmail
    This switch will set the script to send an HTML email report. If this switch is specified, then the To, From and SmtpServers are required.

    .PARAMETER To
    When SendEmail is used, this sets the recipients of the email report.

    .PARAMETER From
    When SendEmail is used, this sets the sender of the email report.

    .PARAMETER SmtpServer
    When SendEmail is used, this is the SMTP Server to send the report through.

    .PARAMETER Subject
    When SendEmail is used, this sets the subject of the email report.

    .PARAMETER NoAttachment
    When SendEmail is used, specifying this switch will set the email report to not include the HTML Report as an attachment. It will still be sent in the body of the email.

#>
param(
    [string[]]$ComputerName = @(),
    [string[]]$FolderPath = @(),
    [switch]$Recurse,
    [switch]$AsHTML,
    [string]$Filename,
    [switch]$SendEmail,
    [string[]]$To,
    [string]$From,
    [string]$SmtpServer,
    [string]$Subject,
    [switch]$NoAttachment
)

# TST 2015-05-26 : measure script execution
$stopWatch = [system.diagnostics.stopwatch]::startNew() 

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

$DateStamp = Get-Date -Format "yyyy-MM-dd"


# Validate parameters
if ($SendEmail)
{
    Write-Verbose "Checking SendEmail requirements"
    
    [array]$newTo = @()
    foreach($recipient in $To)
    {
        if ($recipient -imatch "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z0-9.-]+$")
        {
            $newTo += $recipient
        }
    }
    $To = $newTo
    if (-not $To.Count -gt 0)
    {
        Write-Error "The -To parameter is required when using the -SendEmail switch. If this parameter was used, verify that valid email addresses were specified."
        return
    }
    
    if ($From -inotmatch "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z0-9.-]+$")
    {
        Write-Error "The -From parameter is not valid. This parameter is required when using the -SendEmail switch."
        return
    }

    if ([string]::IsNullOrEmpty($SmtpServer))
    {
        Write-Error "You must specify a SmtpServer. This parameter is required when using the -SendEmail switch."
        return
    }
    if ((Test-Connection $SmtpServer -Quiet -Count 2) -ne $true)
    {
        Write-Error "The SMTP server specified ($SmtpServer) could not be contacted."
        return
    }
}

if (-not $ComputerName.Count -gt 0)
{
    [array]$ComputerName = @()
    Get-ExchangeServer | Where-Object { $_.ServerRole -ilike "*Mailbox*" } | % { $ComputerName += $_.Name }
}

# Build a list of public folders to retrieve
if ($Recurse)
{
    Write-Verbose "Fetching public folder list"
    
    [array]$newFolderPath = @()
    $srvCount = 1
    foreach($srv in $ComputerName)
    {
        $activity = "1: Fetching Public Folder Data"
        $status = "Working on server $($srv)"
        
        Write-Progress -Activity $activity -Status $status -PercentComplete (($srvCount/$ComputerName.Count)*100)
        
        foreach($f in $FolderPath) {
            # ResultSize Unlimited added
            Get-PublicFolder $f -Recurse -ResultSize Unlimited | ForEach-Object { if ($newFolderPath -inotcontains $_.Identity) { $newFolderPath += $_.Identity } }
        }
        $srvCount++
    }
    $FolderPath = $newFolderPath
}

Write-Verbose "Fetching public folder statistics"

# Get statistics for all public folders on all selected servers
# This is significantly faster than trying to get folders one by one by name
[array]$publicFolderList = @()
[array]$nameList = @()
$pfCount = 1
$srvCount = 1
foreach($server in $ComputerName)
{ 
    $pfOnServer = $null
    
    $server = $server.ToUpper()

    $activity = "2: Fetching full public folder statistics"
    $status = "Working on server $($server)"
        
    Write-Progress -Activity $activity -Status $status -PercentComplete (($srvCount/$ComputerName.Count)*100)

    $FileName = "$($DateStamp)-PF-Stats-$($server).xml"
    $File = Join-Path -Path $ScriptDir -ChildPath $FileName
    
    if(Test-Path -Path $File) {
        Write-Progress -Activity $activity -Status "Load stats file $($FileName)" -PercentComplete (($srvCount/$ComputerName.Count)*100)
        $pfOnServer = Import-CliXml -Path $File
    }
    else {
        Write-Progress -Activity $activity -Status $status -PercentComplete (($srvCount/$ComputerName.Count)*100)
        $pfOnServer = Get-PublicFolderStatistics -Server $server -ResultSize Unlimited -ErrorAction SilentlyContinue
        $pfOnServer.FolderPath
    }
    
    if ($FolderPath.Count -gt 0) {
        $pfOnServer = $pfOnServer | Where-Object { $FolderPath -icontains "\$($_.FolderPath)" }
    }
    
    if ($pfOnServer -eq $null) { continue }
    $publicFolderList += New-Object PSObject -Property @{"ComputerName" = $server; "PublicFolderStats" = $pfOnServer}
    $pfOnServer | Foreach-Object { if ($nameList -inotcontains $_.FolderPath) { $nameList += $_.FolderPath } }
    $srvCount++
}
if ($nameList.Count -eq 0)
{
    Write-Error "There are no public folders in the specified servers."
    return
}

$nameList = [array]$nameList | Sort-Object
[array]$ResultMatrix = @()

foreach($folder in $nameList)
{ 
    $resultItem = @{}
    $maxBytes = 0
    $maxSize = $null
    $maxItems = 0
    $srvCount = 1
    foreach($pfServer in $publicFolderList)
    {
        $activity = "3: Checking Public Folder Status"
        $status = "Working on server $($pfServer.ComputerName)"
        
        Write-Progress -Activity $activity -Status $status -PercentComplete (($srvCount/$publicFolderList.Count)*100) 
        
        $pfData = $pfServer.PublicFolderStats | Where-Object { $_.FolderPath -eq $folder }
        if ($pfData -eq $null) { Write-Verbose "Skipping $pfServer.ComputerName for $folder"; continue }
        if (-not $resultItem.ContainsKey("FolderPath"))
        {
            $resultItem.Add("FolderPath", "\$($pfData.FolderPath)")
        }
        if (-not $resultItem.ContainsKey("Name"))
        {
            $resultItem.Add("Name", $pfData.Name)
        }
        if ($resultItem.Data -eq $null)
        {
            $resultItem.Data = @()
        }
        $currentItems = $pfData.ItemCount
        $currentSize = $pfData.TotalItemSize.Value
        
        if ($currentItems -gt $maxItems)
        {
            $maxItems = $currentItems
        }
        if ($currentSize.ToBytes() -gt $maxBytes)
        {
            $maxSize = $currentSize
            $maxBytes = $currentSize.ToBytes()
        }
        $resultItem.Data += New-Object PSObject -Property @{"ComputerName" = $pfServer.ComputerName;"TotalItemSize" = $currentSize; "ItemCount" = $currentItems}
        
        $srvCount++
    }
    $resultItem.Add("TotalItemSize", $maxSize)
    $resultItem.Add("TotalBytes", $maxBytes)
    $resultItem.Add("ItemCount", $maxItems)
    $replCheck = $true
    foreach($dataRecord in $resultItem.Data)
    {
        if ($maxItems -eq 0)
        {
            $progress = 100
        } else {
            $progress = ([Math]::Round($dataRecord.ItemCount / $maxItems * 100, 0))
        }
        if ($progress -lt 100)
        {
            $replCheck = $false
        }
        $dataRecord | Add-Member -MemberType NoteProperty -Name "Progress" -Value $progress
    }
    $resultItem.Add("ReplicationComplete", $replCheck)
    $ResultMatrix += New-Object PSObject -Property $resultItem
    if (-not $AsHTML)
    {
        New-Object PSObject -Property $resultItem
        
    }
}

# TST 2015-05-26 : measure script execution
$stopWatch.Stop()
$elapsedTime = [String]::Format("{0:00}:{1:00}:{2:00}",$stopWatch.Elapsed.Hours,$stopWatch.Elapsed.Minutes,$stopWatch.Elapsed.Seconds)

if ($AsHTML -or $SendEmail -or $Filename -ne $null)
{
    $html = @"
<html>
<style>
body
{
font-family:Arial,sans-serif;
font-size:8pt;
}
table
{
border-collapse:collapse;
font-size:8pt;
font-family:Arial,sans-serif;
border-collapse:collapse;
min-width:400px;
}
table,th, td
{
border: 1px solid black;
}
th
{
text-align:center;
font-size:18;
font-weight:bold;
}
</style>
<body>
<font size="1" face="Arial,sans-serif">
<h1 align="center">Exchange Public Folder Replication Report</h1>
<h4 align="center">Generated $([DateTime]::Now)</h4>
<h5 align="center">Script Runtime $($elapsedTime)</h5>

</font><h2>Overall Summary</h2>
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="2">Public Folder Environment Summary</th></tr>
<tr><td>Servers Selected for this Report</td><td>$($ComputerName -join ", ")</td></tr>
<tr><td>Servers Selected with Public Folders Present</td><td>$(
$serverList = @()
$publicFolderList | ForEach-Object { $serverList += $_.ComputerName }
$serverList -join ", "
)</td></tr>
<tr><td>Number of Public Folders</td><td>$($TotalCount = $ResultMatrix.Count; $TotalCount.ToString("N0"))</td></tr>
<tr><td>Total Size of Public Folders</td><td>$(
$totalSize = $null;
$ResultMatrix | Foreach-Object { $totalSize += $_.TotalItemSize };
$totalSize.ToString("N0")
)</td></tr>
<tr><td>Average Folder Size</td><td>$(($totalSize / $TotalCount))</td></tr>
<tr><td>Total Number of Items in Public Folders</td><td>$(
$totalItemCount = $null
$ResultMatrix | Foreach-Object { $totalItemCount += $_.ItemCount }
$totalItemCount
)</td></tr>
<tr><td>Average Folder Item Count</td><td>$(([Math]::Round($totalItemCount / $TotalCount, 0)).ToString("N0"))</td></tr>
</table>
<br />
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="4">Folders with Incomplete Replication</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td><td>Item Count</td><td>Size</td><td>Servers with Replication Incomplete</td></tr>
$(
[array]$incompleteItems = $ResultMatrix | Where-Object { $_.ReplicationComplete -eq $false }
if (-not $incompleteItems.Count -gt 0)
{
    "<tr><td colspan='4'>There are no public folders with incomplete replication.</td></tr>"
} else {
    foreach($result in $incompleteItems)
    {
        "<tr><td>$($result.FolderPath)</td><td>$($result.ItemCount.ToString("N0"))</td><td>$($result.TotalItemSize.ToString("N0"))</td><td>$(($result.Data | Where-Object { $_.Progress -lt 100 }).ComputerName -join ", ")</td></tr>`r`n"
    }
}
)
</table>
<br />
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="3">Largest Public Folders</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td><td>Item Count</td><td>Size</td></tr>
$(
[array]$largestItems = $ResultMatrix | Sort-Object TotalItemSize -Descending | Select-Object -First 10
if (-not $largestItems.Count -gt 0)
{
    "<tr><td colspan='3'>There are no public folders in this report.</td></tr>"
} else {
    foreach($sizeResult in $largestItems)
    {
        "<tr><td>$($sizeResult.FolderPath)</td><td>$($sizeResult.ItemCount)</td><td>$($sizeResult.TotalItemSize)</td></tr>`r`n"
    }
}
)
</table>

</font><h2>Public Folder Replication Results</h2>
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="$($publicFolderList.Count + 1)">Public Folder Replication Information</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td>
$(
foreach($rServer in $publicFolderList)
{
    "<td>$($rServer.ComputerName)</td>"
}
)
</tr>
$(
if (-not $ResultMatrix.Count -gt 0)
{
    "<tr><td colspan='$($publicFolderList.Count + 1)'>There are no public folders in this report.</td></tr>"
}
foreach($rItem in $ResultMatrix)
{
    "<tr><td>$($rItem.FolderPath)</td>"
    foreach($rServer in $publicFolderList)
    {
        $(
        $rDataItem = $rItem.Data | Where-Object { $_.ComputerName -eq $rServer.ComputerName }
        if ($rDataItem -eq $null)
        {
            "<td>N/A</td>"
        } else {
            if ($rDataItem.Progress -ne 100)
            {
                $color = "#FC2222"
            } else {
                $color = "#A9FFB5"
            }
            "<td style='background-color:$($color)'><div title='$($rDataItem.TotalItemSize) of $($rItem.TotalItemSize) and $($rDataItem.ItemCount) of $($rItem.ItemCount) items.'>$($rDataItem.Progress)%</div></td>"
        }
        )
    }
    "</tr>"
}
)
</table>
</body>
</html>
"@
}

if ($AsHTML)
{
    $html
}

if (-not [string]::IsNullOrEmpty($Filename))
{
    # TST 2015-05-26 : support UTF8 encoding
    $html | Out-File $Filename -Encoding UTF8
}

# TST 2015-05-26 : support UTF8 encoding for Send-MailMessage bug
$utf8 = New-Object System.Text.utf8encoding

if ($SendEmail)
{
    if ([string]::IsNullOrEmpty($Subject))
    {
        $Subject = "Public Folder Environment Report"
    }
    if ($NoAttachment)
    {
        # TST 2015-05-26 : support UTF8 encoding for Send-MailMessage
        Send-MailMessage -SmtpServer $SmtpServer -BodyAsHtml -Body $html -From $From -To $To -Subject $Subject -Encoding $utf8
        
    } else {
        if (-not [string]::IsNullOrEmpty($Filename))
        {
            $attachment = $Filename
        } else {
            $attachment = "$($Env:TEMP)\Public Folder Report - $([DateTime]::Now.ToString("MM-dd-yy")).html"
            $html | Out-File $attachment -Encoding UTF8
        }
        # TST 2015-05-26 : support UTF8 encoding for Send-MailMessage
        Send-MailMessage -SmtpServer $SmtpServer -BodyAsHtml -Body $html -From $From -To $To -Subject $Subject -Attachments $attachment -Encoding $utf8
        Remove-Item $attachment -Confirm:$false -Force
    }
}