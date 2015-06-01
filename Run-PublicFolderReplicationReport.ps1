$publicFolderPath = "\PF1\SUBPF1"

.\Get-PublicFolderReplicationReport.ps1 -ComputerName MXSRV01,MXSRV02,MXSRV03 -FolderPath $publicFolderPath -Recurse -Subject "Public Folder Environment Report [$($publicFolderPath)]" -AsHTML -To thomas@mcsmemail.de -From postmaster@mcsmemail.de -SmtpServer relay.mcsmemail.de -SendEmail