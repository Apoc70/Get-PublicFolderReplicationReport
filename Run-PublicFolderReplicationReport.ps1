param(
    [string]$publicFolderPath = ''
)

# Variables

# Custom label for email subject
$label = 'Exchange 2007'
$recipients = 'pfreports@mcsmemail.de'
$sender = 'postmaster@mcsmemail.de'

# array of public folder servers to query
$publicFolderServers = @('EX2007-01','EX2010-01')

# SMTP server to relay mail
$smtpServer = 'relay.mcsmemail.de'

# Used to trigger a dedicated report for \GrFolder1\Folder1, \GrFolder1\Folder2
$granularRootFolder = @()  # @("\Folder01")
$subPath = ''

# Check for granular folders, Added 2016-01-19
if($granularRootFolder -contains $publicFolderPath) {
    $subPath = $publicFolderPath
    $publicFolderPath = ''
}

#
if($publicFolderPath -ne '') {
    Write-Host "Generating Public Folder reports for $($publicFolderPath)"
    # Generate report for a single public folder | Change COMPUTERNAME attribute for servers to analyse
    .\Get-PublicFolderReplicationReport.ps1 -ComputerName $publicFolderServers -FolderPath $publicFolderPath -Recurse -Subject "Public Folder Environment Report [$($publicFolderPath)] [$($label)]" -AsHTML -To $recipients -From $sender -SmtpServer $smtpServer -SendEmail
}
else {
    if($subPath -ne '') {
        $publicFolderPath = $subPath
    }
    else {
        $publicFolderPath = '\'
    }
    
    if($granularRootFolder.Count -ne 0) {
        Write-Host 'Following root folders will be excluded when using "\":'
        $($granularRootFolder)
    }
    
    Write-Host "Generating Public Folder reports for all folders in $($publicFolderPath)"
    
    $folders = Get-PublicFolder $publicFolderPath -GetChildren 

    # Generate a single report for each folder in root
    $folderCount = ($folders | Measure-Object).Count
    $pfCount = 1
    foreach($pf in $folders) {
        # Check, if folder is in list of granular folders
        if($granularRootFolder -notcontains $pf) {
            if($pf.ParentPath -eq '\') {
                $name = "$($pf.ParentPath)$($pf.Name)"
            }
            else {
                $name = "$($pf.ParentPath)\$($pf.Name)"
            }

            $activity = 'Generating Stats'
            $status = "Fetching $($name)"
            
            Write-Progress -Activity $activity -Status $status -PercentComplete (($pfCount/$folderCount)*100)
          
            .\Get-PublicFolderReplicationReport.ps1 -ComputerName $publicFolderServers -FolderPath $name -Recurse -Subject "Public Folder Environment Report [$($name)] [$($label)]" -AsHTML -To $recipients -From $sender -SmtpServer $smtpServer -SendEmail
            $pfCount++
        }
    }
}