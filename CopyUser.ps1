#Close Chrome
Get-Process -Name "chrome" | Stop-Process -Force -ErrorAction Continue

# Backup user files to a folder named with PC name and username

$pcName = $env:COMPUTERNAME
$userName = $env:USERNAME
$backupRoot = "$PSScriptRoot\Backup"
$backupFolder = Join-Path $backupRoot "$pcName\$userName"

# Chrome Automation 
$proc = 'Chrome'
Start-Process $proc
Start-Sleep 2
$wshell = New-Object -ComObject wscript.shell;
$wshell.AppActivate("$proc")
Start-Sleep 2
$wshell.SendKeys("^{t}")
start-Sleep 3
$wshell.SendKeys('chrome://password-manager/settings')
start-Sleep 3
$wshell.SendKeys("{ENTER}")
start-Sleep 2
$wshell.SendKeys("^{t}")
start-Sleep 3
$wshell.SendKeys('chrome://bookmarks/')
start-Sleep 2
$wshell.SendKeys("{ENTER}")
start-Sleep 2

#cmd /c "rundll32.exe keymgr.dll,KRShowKeyMgr"

$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("After exporting Chrome passwords and bookmarks, click OK to copy the files.", 0, "Export", 0)

# Chrome Automation 
$proc = 'Edge'
Start-Process $proc
Start-Sleep 2
$wshell = New-Object -ComObject wscript.shell;
$wshell.AppActivate("$proc")
Start-Sleep 2
$wshell.SendKeys("^{t}")
start-Sleep 3
$wshell.SendKeys('edge://wallet/passwords?source=assetsSettingsPasswords')
start-Sleep 3
$wshell.SendKeys("{ENTER}")
start-Sleep 2
$wshell.SendKeys("^{t}")
start-Sleep 3
$wshell.SendKeys('edge://favorites/')
start-Sleep 2
$wshell.SendKeys("{ENTER}")
start-Sleep 2

$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("After exporting Edge passwords and bookmarks, click OK to copy the files.", 0, "Export", 0)

cmd /c "rundll32.exe keymgr.dll,KRShowKeyMgr"

$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("After exporting Windows passwords, click OK to copy the files.", 0, "Export", 0)

# List of special folders to back up
$specialFolders = @(
    'Desktop',
    'MyPictures'
)

foreach ($folderName in $specialFolders) {
    $specialFolderEnum = [System.Enum]::Parse([System.Environment+SpecialFolder], $folderName)
    $source = [Environment]::GetFolderPath($specialFolderEnum)
    if ($source -and (Test-Path $source)) {
        $destination = Join-Path $backupFolder (Split-Path $source -Leaf)
        Copy-Item -Path $source -Destination $destination -Recurse -Force
    }
}

#Start-Process explorer.exe -ArgumentList $env:LOCALAPPDATA
$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("Backup complete: $pcName - $userName", 0, "Complete", 0)
Start-Process explorer.exe -ArgumentList $PSScriptRoot\Backup\$pcName\$userName
#Get-Process -Name "cmd" | Stop-Process -Force -ErrorAction Continue
#Write-Host "Backup complete: $backupFolder"
Exit
