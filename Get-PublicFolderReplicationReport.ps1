<#
    .SYNOPSIS
    Generates a report for Exchange Legacy Public Folder Replication.
    
    This is an updated version of the Mike Walker (blog.mikewalker.me) to support non-ASCII environments.

    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
    
    Version 1.6, 2018-01-24

    Ideas, comments and suggestions to support@granikos.eu 
    
    Original Version of the script by Mike Walker: https://gallery.technet.microsoft.com/office/Exchange-2010-Public-944df6ee
    
    .LINK  
    http://scripts.granikos.eu

    .DESCRIPTION
    This script will generate a report for legacy public folder replication. It returns general information, such as total number of public folders, total items in all public folders, total size of all items, the top 10 largest folders, and more. 
    
    Additionally, it lists each Public Folder and the replication status on each server. By default, this script will scan the entire Exchange environment in the current domain and all public folders. This can be limited by using the -ComputerName and -FolderPath parameters.
    
    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1  
    - Exchange Server 2010
    - Exchange Server 2007 (returns sizes as Byte only)

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0 Initial community release of the updated version 
    1.1 Replica status (green/red) depending on item count, not percentage
    1.2 Fixed: If 1st server has a lower item count a folder is not being added to the list of folders with incomplete replication 
    1.3 Changes to number and size formatting  
    1.4 Handling of KB values with Exchange 2007 added 
    1.5 Some PowerShelll hygiene and fixes
    1.6 Count of incomplete replicated public folders stated in table header (issue #1)

    .PARAMETER ComputerName
    This parameter specifies the legacy Exchange server(s) to scan. If this is omitted, all Exchange servers with the Mailbox role in the current domain are scanned.

    .PARAMETER FolderPath
    This parameter specifies the Public Folder(s) to scan. If this is omitted, all public folders are scanned.

    .PARAMETER Recurse
    When used in conjunction with the FolderPath parameter, this will include all child Public Folders of the Folders listed in Folder Path.

    .PARAMETER AsHTML
    Specifying this switch will have this script output HTML, rather than the result objects. This is independent of the Filename or SendEmail parameters and only controls the console output of the script.

    .PARAMETER Filename
    Providing a Filename will save the HTML report to a file. Default = Report.html

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
[CmdletBinding()]
param(
    [string[]]$ComputerName = @(),
    [string[]]$FolderPath = @(),
    [switch]$Recurse,
    [switch]$AsHTML,
    [string]$Filename='Report.html',
    [switch]$SendEmail,
    [string[]]$To,
    [string]$From,
    [string]$SmtpServer,
    [string]$Subject,
    [switch]$NoAttachment
)

# TST 2015-05-26 : measure script execution
$stopWatch = [diagnostics.stopwatch]::startNew() 

$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$DateStamp = Get-Date -Format 'yyyy-MM-dd'
$culture = New-Object -TypeName System.Globalization.CultureInfo -ArgumentList ('de-DE')
$ConvertTo = 'MB' # 'MB', 'GB' !!! Exchange 2007 Only

# TST 2016-05-26 : Convert byte value to MB/GB using the locale settings
function Convert-Value([string]$value) {
    if(((Get-Command -Name exsetup |%{$_.Fileversioninfo}).ProductVersion).StartsWith('08')) {
        # additional calulations for Exchange 2007
        if($value.EndsWith('KB')) {
            # separated each step, instead of a one-liner
            $value = $value.Replace('KB','')
            $vv = ([int]$value) * 1KB
            $value = [string]$vv
        }
        elseif($value.EndsWith('B')) {
            $value = $value.Replace('B','')
        }
        if($value -eq '') {$value = '0'}
        switch($ConvertTo) {
          # Return MB        
          'MB' { $returnValue = "$([math]::round([long]$value/1MB, 2).ToString('n',$culture)) $($ConvertTo)" }
          # Return GB by default
          default { $returnValue = "$([math]::round([long]$value/1GB, 2).ToString('n',$culture)) $($ConvertTo)" }
        }
    } else {
        # keep as default for Exchange 2010
        $returnValue = $value
    }
    return $returnValue
}

$skip = $true
# Validate parameters
if ($SendEmail -and (!$skip)) {
    # Write-Verbose "Checking SendEmail requirements"
    
    [array]$newTo = @()
    foreach($recipient in $To) {
        if ($recipient -imatch "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z0-9.-]+$") {
            $newTo += $recipient
        }
    }
    $To = $newTo
    if (-not $To.Count -gt 0) {
        Write-Error -Message 'The -To parameter is required when using the -SendEmail switch. If this parameter was used, verify that valid email addresses were specified.'
        return
    }
    
    if ($From -inotmatch "^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z0-9.-]+$") {
        Write-Error -Message 'The -From parameter is not valid. This parameter is required when using the -SendEmail switch.'
        return
    }

    if ([string]::IsNullOrEmpty($SmtpServer)) {
        Write-Error -Message 'You must specify a SmtpServer. This parameter is required when using the -SendEmail switch.'
        return
    }
    if ((Test-Connection -ComputerName $SmtpServer -Quiet -Count 2) -ne $true) {
        Write-Error -Message "The SMTP server specified ($SmtpServer) could not be contacted."
        return
    }
}

if (-not $ComputerName.Count -gt 0) {
    [array]$ComputerName = @()
    Get-ExchangeServer | Where-Object { ($_.ServerRole -ilike '*Mailbox*') } | % { $ComputerName += $_.Name } #-and ($_.IsE15OrLater -eq $false)
}

# Build a list of public folders to retrieve
if ($Recurse)
{
    Write-Verbose -Message 'Fetching public folder list'
    
    [array]$newFolderPath = @()
    $srvCount = 1
    foreach($srv in $ComputerName) {
        
        Write-Progress -Activity '1: Fetching Public Folder Data' -Status "Working on server $($srv)" -PercentComplete (($srvCount/$ComputerName.Count)*100)
        
        foreach($f in $FolderPath) {
            # ResultSize Unlimited added
            Get-PublicFolder $f -Recurse -ResultSize Unlimited | ForEach-Object { if ($newFolderPath -inotcontains $_.Identity) { $newFolderPath += $_.Identity } }
        }
        $srvCount++
    }
    $FolderPath = $newFolderPath
}

Write-Verbose -Message "Fetching public folder statistics"

# Get statistics for all public folders on all selected servers
# This is significantly faster than trying to get folders one by one by name
[array]$publicFolderList = @()
[array]$nameList = @()
$pfCount = 1
$srvCount = 1
foreach($server in $ComputerName) { 
    $pfOnServer = $null
    
    $server = $server.ToUpper()

    $activity = '2: Fetching full public folder statistics'
    $status = "Working on server $($server)"
        
    Write-Progress -Activity $activity -Status $status -PercentComplete (($srvCount/$ComputerName.Count)*100)

    $FileNameXml = "$($DateStamp)-PF-Stats-$($server).xml"
    $File = Join-Path -Path $ScriptDir -ChildPath $FileNameXml
    
    if(Test-Path -Path $File) {
        Write-Progress -Activity $activity -Status "Loading stats file $($FileNameXml)" -PercentComplete (($srvCount/$ComputerName.Count)*100)
        $pfOnServer = Import-CliXml -Path $File
    }
    else {
        Write-Progress -Activity $activity -Status $status -PercentComplete (($srvCount/$ComputerName.Count)*100)
        $pfOnServer = Get-PublicFolderStatistics -Server $server -ErrorAction SilentlyContinue -ResultSize Unlimited 
        $pfOnServer.FolderPath
    }
    
    if ($FolderPath.Count -gt 0) {
        $pfOnServer = $pfOnServer | Where-Object { $FolderPath -icontains "\$($_.FolderPath)" }
    }
    
    if ($pfOnServer -eq $null) { continue }
    
    $publicFolderList += New-Object -TypeName PSObject -Property @{"ComputerName" = $server; "PublicFolderStats" = $pfOnServer}
    $pfOnServer | Foreach-Object { if ($nameList -inotcontains $_.FolderPath) { $nameList += $_.FolderPath } }
    $srvCount++
}
if ($nameList.Count -eq 0)
{
    Write-Error -Message "There are no public folders in the specified servers."
    return
}

$nameListMax = $nameList.Count
$nameCount = 1
$nameList = [array]$nameList | Sort-Object
[array]$ResultMatrix = @()

# Check each public folder
foreach($folder in $nameList)
{ 
    $resultItem = @{}
    $maxBytes = 0
    $maxSize = $null
    $maxItems = 0
    $minItems = 0 # 2016-01-15 added, folder incomplete replication fix
    $srvCount = 1
    
    # Check each public folder server in list
    foreach($pfServer in $publicFolderList)
    {
        # 2016-05-25, reordered to display folder name in activity message
        $pfData = $pfServer.PublicFolderStats | Where-Object { $_.FolderPath -eq $folder }
        if ($pfData -eq $null) { 
          Write-Verbose -Message "Skipping $pfServer.ComputerName for $folder"; continue 
        }
        
        $activity = "3: Checking Public Folder Status ($($nameCount)/$($nameListMax)) [\$($pfData.FolderPath)]"
        $status = "Working on server $($pfServer.ComputerName)"
        
        Write-Progress -Activity $activity -Status $status -PercentComplete (($srvCount/$publicFolderList.Count)*100) 
        
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
        # 2016-01-15 added, folder incomplete replication fix
        if (($currentItems -lt $maxItems) -or ($srvCount -eq 1)) {
            $minItems = $currentItems
        }
        
        if ($currentSize.ToBytes() -gt $maxBytes) {
            $maxSize = $currentSize
            $maxBytes = $currentSize.ToBytes()
        }
        $resultItem.Data += New-Object -TypeName PSObject -Property @{
          "ComputerName" = $pfServer.ComputerName
          "TotalItemSize" = $currentSize
          "ItemCount" = $currentItems}
        
        $srvCount++
    }
    
    $resultItem.Add("TotalItemSize", $maxSize)
    $resultItem.Add("TotalBytes", $maxBytes)
    $resultItem.Add("ItemCount", $maxItems)
    $replCheck = $true
    
    foreach($dataRecord in $resultItem.Data) {
        if ($maxItems -eq 0) {
            $progress = 100
        } else {
            $progress = ([Math]::Round($dataRecord.ItemCount / $maxItems * 100, 0))
        }
        if (($progress -lt 100) -or ($minItems -ne $maxItems)) {
            $replCheck = $false
        }
        $dataRecord | Add-Member -MemberType NoteProperty -Name "Progress" -Value $progress
    }
    $resultItem.Add("ReplicationComplete", $replCheck)
    $ResultMatrix += New-Object -TypeName PSObject -Property $resultItem
    if (-not $AsHTML) {
        New-Object -TypeName PSObject -Property $resultItem
    }
    $nameCount++
}

# TST 2015-05-26 : measure script execution
$stopWatch.Stop()
$elapsedTime = [String]::Format("{0:00}:{1:00}:{2:00}",$stopWatch.Elapsed.Hours,$stopWatch.Elapsed.Minutes,$stopWatch.Elapsed.Seconds)

if ($AsHTML -or $SendEmail -or $Filename -ne $null) {
    $activity = "Working..."
    $status = "Generating HTML Report"
        
    Write-Progress -Activity $activity -Status $status -PercentComplete 100
      
    # Html style    
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
<tr><td>Number of Public Folders</td><td>$($TotalCount = $ResultMatrix.Count; $TotalCount.ToString('N0', $culture))</td></tr>
<tr><td>Total Size of Public Folders</td><td>$(
$totalSize = $null
$ResultMatrix | Foreach-Object { $totalSize += $_.TotalItemSize }
Convert-Value -value $totalSize
)</td></tr>
<tr><td>Average Folder Size</td><td>$(Convert-Value -value $($totalSize / $TotalCount) )</td></tr>
<tr><td>Total Number of Items in Public Folders</td><td>$(
$totalItemCount = $null
$ResultMatrix | Foreach-Object { $totalItemCount += $_.ItemCount }
$totalItemCount.ToString('N0', $culture) 
)</td></tr>
<tr><td>Average Folder Item Count</td><td>$( ([Math]::Round($totalItemCount / $TotalCount, 0)).ToString('N0', $culture))</td></tr>
</table>
<br />
<table border="0" cellpadding="3">

$(
[array]$incompleteItems = $ResultMatrix | Where-Object { $_.ReplicationComplete -eq $false }
"<tr style='background-color:#B0B0B0'><th colspan='4'>Folders with Incomplete Replication ($($incompleteItems.Count))</th></tr>
<tr style='background-color:#E9E9E9;font-weight:bold'><td>Folder Path</td><td>Item Count</td><td>Size</td><td>Servers with Replication Incomplete</td></tr>"
if (-not $incompleteItems.Count -gt 0) {
    "<tr><td colspan='4'>There are no public folders with incomplete replication.</td></tr>"
} else {
    foreach($result in $incompleteItems) {
        "<tr><td>$($result.FolderPath)</td><td>$(($result.ItemCount).ToString('N0', $culture))</td><td align='right'>$(Convert-Value -value $($result.TotalItemSize))</td><td>$(($result.Data | Where-Object { $_.Progress -lt 100 }).ComputerName -join ", ")</td></tr>`r`n"
    }
}
)
</table>
<br />
<table border="0" cellpadding="3">
<tr style="background-color:#B0B0B0"><th colspan="3">Largest Public Folders</th></tr>
<tr style="background-color:#E9E9E9;font-weight:bold"><td>Folder Path</td><td>Item Count</td><td>Size</td></tr>
$(
[array]$largestItems = $ResultMatrix | Sort-Object -Property TotalItemSize -Descending | Select-Object -First 20
if (-not $largestItems.Count -gt 0)
{
    "<tr><td colspan='3'>There are no public folders in this report.</td></tr>"
} else {
    foreach($sizeResult in $largestItems)
    {
        "<tr><td>$($sizeResult.FolderPath)</td><td>$(($sizeResult.ItemCount).ToString('N0', $culture))</td><td>$( Convert-Value -value $($sizeResult.TotalItemSize))</td></tr>`r`n"
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
        if ($rDataItem -eq $null) {
            "<td>N/A</td>"
        } else {
            if ($rDataItem.ItemCount -ne $rItem.ItemCount) { #$rDataItem.Progress -ne 100 {
                $color = "#FC2222"
            } else {
                $color = "#A9FFB5"
            }
            "<td style='background-color:$($color)'><div title='$(Convert-Value -value $($rDataItem.TotalItemSize)) of $(Convert-Value -value $($rItem.TotalItemSize)) and $(($rDataItem.ItemCount).ToString('N0', $culture)) of $(($rItem.ItemCount).ToString('N0', $culture)) items.'>$($rDataItem.Progress)% [$(($rDataItem.ItemCount).ToString('N0', $culture))]</div></td>"
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

if ($AsHTML) {
    $html
}

if (-not [string]::IsNullOrEmpty($Filename)) {
    # TST 2015-05-26 : support UTF8 encoding
    $html | Out-File -FilePath $Filename -Encoding UTF8
}

# TST 2015-05-26 : support UTF8 encoding for Send-MailMessage bug
$utf8 = New-Object -TypeName System.Text.utf8encoding

if ($SendEmail) {
    if ([string]::IsNullOrEmpty($Subject)) {
        $Subject = "Public Folder Environment Report"
    }
    if ($NoAttachment) {
        # TST 2015-05-26 : support UTF8 encoding for Send-MailMessage
        Send-MailMessage -SmtpServer $SmtpServer -BodyAsHtml -Body $html -From $From -To $To -Subject $Subject -Encoding $utf8
        
    } else {
        if (-not [string]::IsNullOrEmpty($Filename)) {
            $attachment = $Filename
        } else {
            $attachment = "$($Env:TEMP)\Public Folder Report - $([DateTime]::Now.ToString("MM-dd-yy")).html"
            $html | Out-File -FilePath $attachment -Encoding UTF8
        }
        # TST 2015-05-26 : support UTF8 encoding for Send-MailMessage
        Send-MailMessage -SmtpServer $SmtpServer -BodyAsHtml -Body $html -From $From -To $To -Subject $Subject -Attachments $attachment -Encoding $utf8
        Remove-Item -Path $attachment -Confirm:$false -Force
    }
}