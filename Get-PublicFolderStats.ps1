param(
    [string[]]$ComputerName = @(),
    [switch]$Reload
)

if (-not $ComputerName.Count -gt 0)
{

   $ComputerName = @("SERVER01","SERVER02","SERVER03")
}

$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path

$DateStamp = Get-Date -Format "yyyy-MM-dd"

foreach($server in $ComputerName) {

    $server = $server.ToUpper()

    $StatsFileName = "$($DateStamp)-PF-Stats-$($server).xml"
    
    $StatsFile = Join-Path -Path $ScriptDir -ChildPath $StatsFileName
    
    Write-Host "Saving public folder stats to: $StatsFile"
    
    if(!(Test-Path $StatsFile) -or $Reload) {
        Write-Host "> File does not exist! Load stats from $($server)" 
        # $stats = Get-PublicFolder "\" -Recurse | Get-PublicFolderStatistics -Server $server
        $stats = Get-PublicFolderStatistics -Server $server -ResultSize Unlimited -ErrorAction SilentlyContinue
        $stats | Export-Clixml -Path $StatsFile -Force
    }
    else {
        Write-Host "> File exists, nothing to do!"
    }
}

foreach($server in $ComputerName) {

    $server = $server.ToUpper()

    $PubFileName = "$($DateStamp)-PF-Data-$($server).xml"
    
    $PubFile = Join-Path -Path $ScriptDir -ChildPath $PubFileName
    
    Write-Host "Saving public data: $PubFile"
    
    if(!(Test-Path $PubFile) -or $Reload) {
        Write-Host "> File does not exist! Load public folder from $($server)" 
        $pub = Get-PublicFolder -Server $server -Recurse -ResultSize Unlimited -ErrorAction SilentlyContinue
        $pub | Export-Clixml -Path $PubFile -Force -Encoding UTF8
    }
    else {
        Write-Host "> File exists, nothing to do!"
    }
}