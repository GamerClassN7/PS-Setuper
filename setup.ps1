param(
    [Parameter(Mandatory = $false)]
    [Switch]$recovery
)

#Administrator Permission Level Check
$isAdministrator = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not ($isAdministrator)) {
    Write-Warning "Script need to be run as administrator !!!"
    # return
}

#Get Enviroment Variables
$osName = (Get-ComputerInfo -Property OsName).OsName;
$userName = $env:UserName;
$rootBackupPath = "C:\Users\{0}\AppData\Local\Temp\setuperBackup" -f $userName;
New-Item -Path ("C:\Users\{0}\AppData\Local\Temp\" -f $userName) -Name "setuperBackup" -ItemType "directory" -Force


#Menu & Intro
write-host ("Hello '{0}'" -f $userName)
write-host ("OS '{0}'" -f $osName)

#Backup FIlezilla Configs
$fzBackupPath = $rootBackupPath +"\fzExport\";
New-Item -Path $rootBackupPath -Name "fzExport" -ItemType "directory" -Force
Copy-Item ("C:\Users\{0}\AppData\Roaming\FileZilla\filezilla.xml" -f $userName) -Destination $fzBackupPath -Force
Copy-Item ("C:\Users\{0}\AppData\Roaming\FileZilla\layout.xml" -f $userName) -Destination $fzBackupPath -Force
Copy-Item ("C:\Users\{0}\AppData\Roaming\FileZilla\sitemanager.xml" -f $userName) -Destination $fzBackupPath -Force

#Backup Tabby Configs
$tbBackupPath = $rootBackupPath +"\tbExport\";
New-Item -Path $rootBackupPath -Name "tbExport" -ItemType "directory" -Force
Copy-Item ("C:\Users\{0}\AppData\Roaming\tabby\config.yaml" -f $userName) -Destination $tbBackupPath -Force

#Backup KeeWeb Configs
$tbBackupPath = $rootBackupPath +"\kwExport\";
New-Item -Path $rootBackupPath -Name "kwExport" -ItemType "directory" -Force
Copy-Item ("C:\Users\{0}\AppData\Roaming\KeeWeb\app-settings.dat" -f $userName) -Destination $tbBackupPath -Force

function New-RegFolder {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $FolderName,
        [Parameter()]
        [string]
        $FolderPath,
        [Parameter()]
        [string]
        $DefaultValue = $null
    )

    $regExist = ((Get-Item -Path ("{0}{1}" -f $FolderPath, $FolderName) -ErrorAction SilentlyContinue | Measure-Object).Count -gt 0)
    if (-not $regExist) {
        $newKey = @{
            Path = $FolderPath
            Name = $FolderName
        }

        if ($DefaultValue -ne $null) {
            $newKey["Value"] = $DefaultValue
        }

        New-Item @newKey
    }
    else {
        Write-Host "exist" + $FolderPath + $FolderName
    }
}

function New-RegItem {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $ItemName,
        [Parameter()]
        [string]
        $ItemPath,
        [Parameter()]
        [string]
        $ItemValue,
        [Parameter()]
        [string]
        $ItemType = "DWORD"
    )

    $regItemExist = (((Get-ItemProperty -Path ("{0}" -f $ItemPath) -ErrorAction SilentlyContinue).PSObject.Properties | Where-Object -Filter { $_.Name -eq $ItemName } | Measure-Object).Count -gt 0)
    if (-not $regItemExist) {
        $newItem = @{
            Path  = $ItemPath
            Name  = $ItemName
            Value = $ItemValue
            Type  = $ItemType
        }

        New-ItemProperty @newItem
    }
    else {
        Write-Host "exist" + $ItemPath + "[" + $ItemType + "]"+ $ItemName + ":" + $ItemValue
    }
}

function Set-RegItem {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $ItemName,
        [Parameter()]
        [string]
        $ItemPath,
        [Parameter()]
        [string]
        $ItemValue,
        [Parameter()]
        [string]
        $ItemType = "DWORD"
    )

    $newItem = @{
        Path  = $ItemPath
        Name  = $ItemName
        Value = $ItemValue
        Type  = $ItemType
    }

    Set-ItemProperty @newItem
}

function Write-ZipUsing7Zip([string]$FilesToZip, [string]$ZipOutputFilePath, [string]$Password, [ValidateSet('7z','zip','gzip','bzip2','tar','iso','udf')][string]$CompressionType = 'zip', [switch]$HideWindow)
{
    # Look for the 7zip executable.
    $pathTo32Bit7Zip = "C:\Program Files (x86)\7-Zip\7z.exe"
    $pathTo64Bit7Zip = "C:\Program Files\7-Zip\7z.exe"
    $THIS_SCRIPTS_DIRECTORY = Split-Path $script:MyInvocation.MyCommand.Path
    $pathToStandAloneExe = Join-Path $THIS_SCRIPTS_DIRECTORY "7za.exe"
    if (Test-Path $pathTo64Bit7Zip) { $pathTo7ZipExe = $pathTo64Bit7Zip }
    elseif (Test-Path $pathTo32Bit7Zip) { $pathTo7ZipExe = $pathTo32Bit7Zip }
    elseif (Test-Path $pathToStandAloneExe) { $pathTo7ZipExe = $pathToStandAloneExe }
    else { throw "Could not find the 7-zip executable." }

    # Delete the destination zip file if it already exists (i.e. overwrite it).
    if (Test-Path $ZipOutputFilePath) { Remove-Item $ZipOutputFilePath -Force }

    $windowStyle = "Normal"
    if ($HideWindow) { $windowStyle = "Hidden" }

    # Create the arguments to use to zip up the files.
    # Command-line argument syntax can be found at: http://www.dotnetperls.com/7-zip-examples
    $arguments = "a -t$CompressionType ""$ZipOutputFilePath"" ""$FilesToZip"" -mx9"
    if (!([string]::IsNullOrEmpty($Password))) { $arguments += " -p$Password" }

    # Zip up the files.
    $p = Start-Process $pathTo7ZipExe -ArgumentList $arguments -Wait -PassThru -WindowStyle $windowStyle

    # If the files were not zipped successfully.
    if (!(($p.HasExited -eq $true) -and ($p.ExitCode -eq 0)))
    {
        throw "There was a problem creating the zip file '$ZipFilePath'."
    }
}

#OS Common
if ($osName.startsWith("Microsoft Windows 10") -or $osName.startsWith("Microsoft Windows 11")) {
    #Dark-Mode
    Set-RegItem -ItemPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -ItemName "AppsUseLightTheme" -ItemType "DWORD" -ItemValue 0
    Set-RegItem -ItemPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -ItemName "ColorPrevalence" -ItemType "DWORD" -ItemValue 0
    
    #File Extensions
    Set-RegItem -ItemPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -ItemName "HideFileExt" -ItemType "DWORD" -ItemValue 0

    #Visueal Efects
    Set-RegItem -ItemPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects" -ItemName "VisualFXSetting" -ItemType "DWORD" -ItemValue 4
}

#OS Specific Tasks
if ($osName.startsWith("Microsoft Windows 10")) {
    #Unistall One Drive
    Stop-Process -Force -Name "OneDrive.exe"
    Start-Process "$env:windir\SysWOW64\OneDriveSetup.exe" "/uninstall"
}

if ($osName.startsWith("Microsoft Windows 11")) {
    #Old Context Menu
    New-RegFolder -FolderPath "HKCU:\Software\Classes\CLSID\" -FolderName "{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" 
    New-RegFolder -FolderPath "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\" -FolderName "InprocServer32" -DefaultValue ""

    #Start Menu to left
    New-RegItem -ItemPath "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -ItemName "TaskbarAl" -ItemType "DWORD" -ItemValue 0
}
    
# <# Create Data Partition
# if one disk select first
#  $freeDiskSpaceGb = ([math]::floor((Get-Volume -DriveLetter C).SizeRemaining /1GB) - 1)
#  if ($freeDiskSpaceGb -gt 10){
#      New-Partition -DiskNumber 0 -Size 10GB -DriveLetter V | Format-Volume -FileSystem NTFS -NewFileSystemLabel 'WORK 2'
#  } 
    
# Add local Admin Creation
# Add disable of Administraor Default #>



#NEW ITERATION#
##Backup Wifis##
$wifiBackupPath = $rootBackupPath +"\wfExport";
if ($recovery.IsPresent) {
    Get-ChildItem $wifiBackupPath | ForEach-Object {
        $profilePath = ($wifiBackupPath + "\" + $_.Name);
        % { (netsh wlan add profile file="$profilePath" user=current) } 
    } 
}
else {
    New-Item -Path $rootBackupPath -Name "wfExport" -ItemType "directory" -Force
    (netsh wlan show profiles) | Select-String "\:(.+)$" | % { $name = $_.Matches.Groups[1].Value.Trim(); $_ } | % { (netsh wlan export profile name="$name" folder="$wifiBackupPath" key=clear) }
}

##WINGET##
$hasPackageManager = Get-AppPackage -name "Microsoft.DesktopAppInstaller"
if (!$hasPackageManager) {
    $releases_url = "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $releases = Invoke-RestMethod -uri "$($releases_url)"
    $latestRelease = $releases.assets | Where-Object { $_.browser_download_url.EndsWith("msixbundle") } | Select-Object -First 1
    Add-AppxPackage -Path $latestRelease.browser_download_url
}
write-host ("Winget Installed")
    
$wingetBackupPath = $rootBackupPath +"\wgExport.json";
if ($recovery.IsPresent) {
    if (-not [System.IO.File]::Exists($wingetBackupPath)) {

        $wingetAppList = Get-Content $wingetBackupPath | ConvertFrom-Json
        $wingetAppList.Sources.Packages | ForEach-Object {
            #TODO: Speed Up Iteration
            $exist = (winget list $_.PackageIdentifier).Trim()
            if ($exist -eq "No installed package found matching input criteria.") {
                write-host $_.PackageIdentifier
                winget install --id $_.PackageIdentifier
                return;
            }
            write-host ("Already installed '{0}'" -f $_.PackageIdentifier)
        }
    }
    else {
        write-host ("no VS code backup file found")
    }
}
else {
    winget export -o "$wingetBackupPath" >> $null
}

##VS CODE##
$vsBackupPath = $rootBackupPath +"\vsExport.json";
if ($recovery.IsPresent) {
    if (-not [System.IO.File]::Exists($vsBackupPath)) {
        $vsExtensionsList = Get-Content $vsBackupPath | ConvertFrom-Json
        $vsExtensionsListActual = & code --list-extensions
        $vsExtensionsList | Where-Object -Filter { -not $vsExtensionsListActual.Contains($_) } | ForEach-Object {
            write-host $_
            code --install-extension $_
        }
    }
    else {
        write-host ("no VS code backup file found")
    }
}
else {
    $commandOutput = & code --list-extensions
    $commandOutput.Split([Environment]::NewLine) | ConvertTo-Json | Out-File -Encoding utf8 -FilePath $vsBackupPath
}


Write-ZipUsing7Zip -FilesToZip $rootBackupPath -ZipOutputFilePath ("C:\Users\{0}\Desktop\SetuperBackup.zip" -f $userName) -Password "password123"

Write-Host "DONE :)"

