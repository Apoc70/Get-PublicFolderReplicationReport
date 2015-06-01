# Get-PublicFolderReplicationReport.ps1
This is an updated version of the script originally posted by Mike Walker here: https://gallery.technet.microsoft.com/office/Exchange-2010-Public-944df6ee

The updated version of script allows for reports showing public folder names having *non-ASCII* characters.

Simplify the script execution by creating a simple Run-PublicFolderReplicationReport.ps1

$publicFolderPath = "\PF1\SUBPF1"
.\Get-PublicFolderReplicationReport.ps1 -ComputerName MXSRV01,MXSRV02,MXSRV03 -FolderPath $publicFolderPath -Recurse -Subject "Public Folder Environment Report [$($publicFolderPath)]" -AsHTML -To thomas@mcsmemail.de -From postmaster@mcsmemail.de -SmtpServer relay.mcsmemail.de -SendEmail

