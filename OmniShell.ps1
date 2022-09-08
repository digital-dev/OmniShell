function listarchives {
  $outlook = New-Object -comObject Outlook.Application
  $outlook.Session.Stores | where { ($_.FilePath -like '*.PST') } | format-table DisplayName, FilePath -autosize
}
function getspace($pcname) {
  $disk = Get-WmiObject Win32_LogicalDisk -ComputerName $pcname -Filter "DeviceID='C:'" |
  Select-Object Size,FreeSpace
  $disk.Size / 1GB
  $disk.FreeSpace / 1GB
}
function dropshell($pcname) {
  $UserCredential = Get-Credential
  $sess = New-PSSession -ComputerName $pcname -Credential $UserCredential
  Enter-PSSession $sess
}
function connect365 {
  $UserCredential = Get-Credential
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Kerberos 
}
function connectExchange {
  $UserCredential = Get-Credential
  $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://<ServerFQDN>/PowerShell/ -Authentication Kerberos -Credential $UserCredential
}
function cleanup($pcname) {
  $UserCredential = Get-Credential
  Invoke-Command -ComputerName $pcname -Credential $UserCredential -ScriptBlock {
  function global:Write-Verbose ( [string]$Message ) {
  # Check $VerbosePreference variable, and turns -Verbose on
  if ( $VerbosePreference -ne 'SilentlyContinue' )
  { Write-Host " $Message" -ForegroundColor 'Yellow' }}
  $VerbosePreference = "Continue"
  $DaysToDelete = 1
  $LogDate = get-date -format "MM-d-yy-HH"
  $objShell = New-Object -ComObject Shell.Application 
  $objFolder = $objShell.Namespace(0xA)
  $ErrorActionPreference = "silentlycontinue"
  Start-Transcript -Path C:\Cleanup-$LogDate.log
  ## Cleans all code off of the screen.
  Clear-Host
  $size = Get-ChildItem C:\Users\* -Include *.iso, *.vhd, *.vmdk -Recurse -ErrorAction SilentlyContinue | 
  Sort Length -Descending | 
  Select-Object Name,
  @{Name="Size (GB)";Expression={ "{0:N2}" -f ($_.Length / 1GB) }}, Directory |
  Format-Table -AutoSize | Out-String
  $Before = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq "3" } | Select-Object SystemName,
  @{ Name = "Drive" ; Expression = { ( $_.DeviceID ) } },
  @{ Name = "Size (GB)" ; Expression = {"{0:N1}" -f( $_.Size / 1gb)}},
  @{ Name = "FreeSpace (GB)" ; Expression = {"{0:N1}" -f( $_.Freespace / 1gb ) } },
  @{ Name = "PercentFree" ; Expression = {"{0:P1}" -f( $_.FreeSpace / $_.Size ) } } |
  Format-Table -AutoSize | Out-String                                  
  ## Stops the windows update service. 
  Get-Service -Name wuauserv | Stop-Service -Force -Verbose -ErrorAction SilentlyContinue
  ## Windows Update Service has been stopped successfully!
  ## Deletes the contents of windows software distribution.
  Get-ChildItem "C:\Windows\SoftwareDistribution\*" -Recurse -Force -Verbose -ErrorAction SilentlyContinue | remove-item -force -Verbose -recurse -ErrorAction SilentlyContinue
  ## The Contents of Windows SoftwareDistribution have been removed successfully!
  ## Deletes the contents of the Windows Temp folder.
  Get-ChildItem "C:\Windows\Temp\*" -Recurse -Force -Verbose -ErrorAction SilentlyContinue |
  Where-Object { ($_.CreationTime -lt $(Get-Date).AddDays(-$DaysToDelete)) } |
  remove-item -force -Verbose -recurse -ErrorAction SilentlyContinue
  ## The Contents of Windows Temp have been removed successfully!
  ## Delets all files and folders in user's Temp folder. 
  Get-ChildItem "C:\users\*\AppData\Local\Temp\*" -Recurse -Force -ErrorAction SilentlyContinue |
  Where-Object { ($_.CreationTime -lt $(Get-Date).AddDays(-$DaysToDelete))} |
  remove-item -force -Verbose -recurse -ErrorAction SilentlyContinue
  ## The contents of C:\users\$env:USERNAME\AppData\Local\Temp\ have been removed successfully!
  ## Remove all files and folders in user's Temporary Internet Files. 
  Get-ChildItem "C:\users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*" `
  -Recurse -Force -Verbose -ErrorAction SilentlyContinue |
  Where-Object {($_.CreationTime -le $(Get-Date).AddDays(-$DaysToDelete))} |
  remove-item -force -recurse -ErrorAction SilentlyContinue
  ## All Temporary Internet Files have been removed successfully!
  ## Cleans IIS Logs if applicable.
  Get-ChildItem "C:\inetpub\logs\LogFiles\*" -Recurse -Force -ErrorAction SilentlyContinue |
  Where-Object { ($_.CreationTime -le $(Get-Date).AddDays(-60)) } |
  Remove-Item -Force -Verbose -Recurse -ErrorAction SilentlyContinue
  ## All IIS Logfiles over x days old have been removed Successfully!
  ## Deletes the contents of the recycling Bin.
  Clear-RecycleBin -Force
  $objFolder.items() | ForEach-Object { Remove-Item $_.path -ErrorAction Ignore -Force -Verbose -Recurse }
  ## Starts the Windows Update Service
  ##Get-Service -Name wuauserv | Start-Service -Verbose
  $After =  Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq "3" } | Select-Object SystemName,
  @{ Name = "Drive" ; Expression = { ( $_.DeviceID ) } },
  @{ Name = "Size (GB)" ; Expression = {"{0:N1}" -f( $_.Size / 1gb)}},
  @{ Name = "FreeSpace (GB)" ; Expression = {"{0:N1}" -f( $_.Freespace / 1gb ) } },
  @{ Name = "PercentFree" ; Expression = {"{0:P1}" -f( $_.FreeSpace / $_.Size ) } } |
  Format-Table -AutoSize | Out-String
  ## Sends some before and after info for ticketing purposes
  Hostname ; Get-Date | Select-Object DateTime
  Write-Verbose "Before: $Before"
  Write-Verbose "After: $After"
  Write-Verbose $size
  Stop-Transcript}
}
clear
Write-Host "getspace <pc_name> : Gets the amount of free disk space on remote computer."
Write-Host "cleanup <pc_name> : Cleans disk space for remote computer."
Write-Host "dropshell <pc_name> : Drops into remote powershell session."
Write-Host "connect365 : Connects to the CLI interface for Office 365 Administration."
Write-Host "connectExchange :  Connects to the CLI interface for Exchange Administration."
# Seperate Shell Maybe? (Local import onto remote user PC as their user for mailbox recreation automation)