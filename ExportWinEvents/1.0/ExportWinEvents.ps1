#========================================================================
# Created with: SAPIEN Technologies, Inc., PowerShell Studio 2012 v3.1.35
# Created on:   2/26/2022 9:23 PM
# Created by:   Inderjeet
# Filename	: 	ExportWinEvents.PS1     
#========================================================================

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path
[xml]$ConfigFile = Get-Content "ExportEventSettings.xml"
$logArray = @()
Foreach($log in $ConfigFile.configuration.settings)
{
    $logArray = $log.Property.value
}


# Logs to extract from server
Import-Module AppvClient
#$logArray = @("System","Security","Application")

# Grabs the server name to append to the log file extraction
$servername = $env:computername

# Provide the path with ending "\" to store the log file extraction.
#$destinationpath = "C:\ArchiveEventLogs\"
$destinationpath = $ConfigFile.configuration.path.property.value

If((Test-Path $destinationpath) -eq $false)
{
    New-Item -Name "ArchiveEventLogs" -ItemType Directory -Path "C:\"
}

# changed the logic of getting the number of days to keep the zipped files
#$limit = (Get-Date).AddDays(-15)

[int]$duration = [int]$ConfigFile.configuration.duration.days
$limit = (Get-Date).AddDays($duration)

# Delete files older than the $limit.
Get-ChildItem -Path $destinationpath -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force

# Delete any empty directories left behind after deleting the old files.
Get-ChildItem -Path $destinationpath -Recurse -Force | Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } | Remove-Item -Force -Recurse

# Checks the last character of the destination path.  If it does not end in '\' it adds one.
# '.+?\\$' +? means any character \\ is looking for the backslash $ is the end of the line charater
if ($destinationpath -notmatch '.+?\\$')
{
    $destinationpath += '\'
}

# If the destination path does not exist it will create it
if (!(Test-Path -Path $destinationpath))
{
    New-Item -ItemType directory -Path $destinationpath
}

$FolderPath = $destinationpath + (Get-Date).tostring("MM-dd-yyyy")

If((Test-Path $FolderPath) -eq $false)
{
	New-Item -Path $destinationpath -Name (Get-Date).tostring("MM-dd-yyyy") -ItemType "directory"
}

# Get the current date in YearMonthDay format
$logdate = Get-Date -format yyyyMMddHHmm

# Start Process Timer
$StopWatch = [system.diagnostics.stopwatch]::startNew()

Foreach($log in $logArray)
{
    # If using Clear and backup
    $destination = $FolderPath + "\" + $log + ".evtx"

    Write-Host "Extracting the $log file now."

    # Extract each log file listed in $logArray from the local server.
    wevtutil epl $log $destination

    Write-Host "Clearing the $log file now."

    # Clear the log and backup to file.
    WevtUtil cl $log

}

# End Code

# Stop Timer
$StopWatch.Stop()
$TotalTime = $StopWatch.Elapsed.TotalSeconds
$TotalTime = [math]::Round($totalTime, 2)
write-host "The Script took $TotalTime seconds to execute."

$ZIPpath = $destinationpath + (Get-Date).tostring("MM-dd-yyyy") + ".ZIP"
If((Test-Path $ZIPpath) -eq $true)
{
	$Num = Get-Random -Maximum 100 
	$ZIPpath = $destinationpath + (Get-Date).tostring("MM-dd-yyyy") + "-" + $Num +".ZIP"
}
Compress-Archive -Path $FolderPath -DestinationPath $ZIPpath

Get-ChildItem -Path $FolderPath -Recurse -Force | Remove-Item -Force -Recurse
Remove-Item $FolderPath -Force -Recurse
