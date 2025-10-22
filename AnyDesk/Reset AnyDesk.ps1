# Check if the script is running as administrator
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
	Write-Host "Please run as administrator."
	Pause
	exit
}

Write-Host "Restarting AnyDesk..." -ForegroundColor Cyan

function Stop-AnyDesk {
	Stop-Service -Name "AnyDesk" -ErrorAction SilentlyContinue
	Stop-Process -Name "AnyDesk" -Force -ErrorAction SilentlyContinue
	Start-Sleep -Seconds 1
}

function Start-AnyDesk {
	Start-Service -Name "AnyDesk" -ErrorAction SilentlyContinue
	$paths = @("$env:ProgramFiles\AnyDesk\AnyDesk.exe", "$env:ProgramFiles(x86)\AnyDesk\AnyDesk.exe")
	foreach ($path in $paths) {
		if (Test-Path $path) {
			Start-Process $path
			break
		}
	}
}

# Stop AnyDesk
Stop-AnyDesk

# Paths
$allUsers = "$env:ProgramData\AnyDesk"
$appData = "$env:AppData\AnyDesk"
$temp = "$env:TEMP"

# Remove configuration files
Remove-Item "$allUsers\service.conf" -Force -ErrorAction SilentlyContinue
Remove-Item "$appData\service.conf" -Force -ErrorAction SilentlyContinue

# Backup current user.conf
Copy-Item "$appData\user.conf" -Destination "$temp\" -Force -ErrorAction SilentlyContinue

# Thumbnails
Remove-Item "$temp\thumbnails" -Recurse -Force -ErrorAction SilentlyContinue
Copy-Item "$appData\thumbnails" -Destination "$temp\thumbnails" -Recurse -Force -ErrorAction SilentlyContinue

# Clean folders
Remove-Item "$allUsers\*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$appData\*" -Recurse -Force -ErrorAction SilentlyContinue

# Restart the service
Start-AnyDesk

# Wait for system.conf with ID to be created
while (-not (Select-String -Path "$allUsers\system.conf" -Pattern "ad.anynet.id=" -Quiet)) {
	Start-Sleep -Seconds 1
}

# Restore settings
Stop-AnyDesk
Move-Item "$temp\user.conf" "$appData\user.conf" -Force -ErrorAction SilentlyContinue
Copy-Item "$temp\thumbnails" -Destination "$appData\thumbnails" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item "$temp\thumbnails" -Recurse -Force -ErrorAction SilentlyContinue

Start-AnyDesk

Write-Host "*********" -ForegroundColor Green
Write-Host "Done." -ForegroundColor Green
Pause
