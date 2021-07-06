Write-Output "Setting up Teams Virtual Users..."
Write-Output "Installing Chocolatey, OBS and Teams"
# Install Chocolatey
Set-ExecutionPolicy Bypass -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
# Install Applications
choco install obs-studio.install -y
choco install microsoft-teams -y
Write-Output "Downloading Custom OBS Configuration"
$BaseURL="https://github.com/spgoodman/TeamsVirtualUsers/raw/main"
Write-Output "Downloading OBS Profile ZIP file"
Invoke-WebRequest -Uri "$($BaseURL)/obs-studio-settings-profile-scenes.zip" -OutFile "$($Env:APPDATA)\obs-studio-settings-profile-scenes.zip" -UseBasicParsing
if (Test-Path "$($Env:APPDATA)\obs-studio") {
    $Answer = Read-Host -Prompt "OBS Profile Folder Exists. Enter Y to overwrite"
    if ($Answer -ieq "y")
    {
        Expand-Archive "$($Env:APPDATA)\obs-studio-settings-profile-scenes.zip" -DestinationPath "$($Env:APPDATA)" -Force 
    } else {
        break
    }
} else {
    Expand-Archive "$($Env:APPDATA)\obs-studio-settings-profile-scenes.zip" -DestinationPath "$($Env:APPDATA)"
}
Write-Output "Downloading Virtual User Video Files"
$VirtualUsers=@("Chandler","Joey","Monica","Phoebe","Rachael","Ross","Black Panther","Iron Man","Star Lord","Thanos","Thor","Darth Vader")
$VirtualUserDir = "$($Env:USERPROFILE)\Videos\Teams Virtual Users"
if (!(Test-Path $VirtualUserDir))
{
    New-Item -Path $VirtualUserDir -ItemType Directory
}
foreach ($VirtualUser in $VirtualUsers)
{
    Invoke-WebRequest -Uri "$($BaseURL)/$($VirtualUser).mp4" -OutFile "$($VirtualUserDir)\$($VirtualUser).mp4" -UseBasicParsing
}
Write-Output "Updating OBS Scenes"
$VirtualUsersScenes = Get-Content -Raw -Path "$($Env:APPDATA)\obs-studio\basic\scenes\VirtualUsers.json"
$VirtualUsersScenes = $VirtualUsersScenes -replace "C:/users/admin",$env:UserProfile -replace "\\","/"
$VirtualUsersScenes | Out-File -FilePath "$($Env:APPDATA)\obs-studio\basic\scenes\VirtualUsers.json"
Write-Output "Choose the character to show when OBS launches (you can change this later in the app)"
for ($i=0; $i -lt $VirtualUsers.Count; $i++)
{
    Write-Output " [$($i+1)] - $($VirtualUsers[$i])"
}
$Choice=0
while (!($Choice -in 1 .. $VirtualUsers.Count))
{
    $Choice = Read-Host -Prompt "Choose [1-$($VirtualUsers.Count)]"
    $Choice
}
Write-Output "You chose $($VirtualUsers[$Choice-1]) - updating scene file with current scene selection"
$VirtualUsersScenes = $VirtualUsersScenes -replace "`"current_program_scene`":`"Chandler`",`"current_scene`":`"Chandler`"","`"current_program_scene`":`"$($VirtualUsers[$Choice-1])`",`"current_scene`":`"$($VirtualUsers[$Choice-1])`""
$VirtualUsersScenes | Out-File -FilePath "$($Env:APPDATA)\obs-studio\basic\scenes\VirtualUsers.json" -Encoding UTF8

Write-Output "Creating OBS as a Startup App with Virtual Webcam Enabled"
$ShortcutFile = "$($Env:APPDATA)\Microsoft\Windows\Start Menu\Programs\Startup\OBS Studio - Virtual Webcam Enabled.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = "C:\Program Files\obs-studio\bin\64bit\obs64.exe"
$Shortcut.Arguments = "--startvirtualcam"
$Shortcut.WorkingDirectory = "C:\Program Files\obs-studio\bin\64bit\"
$Shortcut.Save()

Write-Output "Optionally setting Auto Login for the Windows PC"
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
$Password = Read-Host -Prompt "Enter this Windows user's login password to enable automatic login, or hit enter to skip" 
if ($Password)
{
    Write-Output "Setting default username ($($env:username) and password for AutoAdminLogin"
    Set-ItemProperty $RegPath "AutoAdminLogon" -Value "1" -type String 
    Set-ItemProperty $RegPath "DefaultUsername" -Value $env:username -type String 
    Set-ItemProperty $RegPath "DefaultPassword" -Value $Password -type String
}
