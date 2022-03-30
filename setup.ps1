

#Administrator Permission Level Check
$isAdministrator = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not ($isAdministrator)) {
    Write-Warning "Script need to be run as administrator !!!"
    return
}

#Get Enviroment Variables
$osName = (Get-ComputerInfo -Property OsName).OsName;
$userName = $env:UserName;

#Menu & Intro
write-host ("Hello '{0}'" -f $userName)

#WinGet
$wingetBackupPath = "C:\Users\{0}\AppData\Local\Temp\wgExport.json" -f $userName;
if (-not [System.IO.File]::Exists($wingetBackupPath)){
    winget export -o "$wingetBackupPath" >> $null
}

$wingetAppList = Get-Content $wingetBackupPath | ConvertFrom-Json
$wingetAppList.Sources.Packages | ForEach-Object {
    #TODO: Speed Up Iteration
    winget list $_.PackageIdentifier >> $exist
    if ($exist -eq "No installed package found matching input criteria."){
        write-host $_.PackageIdentifier
        winget install --id="$_.PackageIdentifier"
        return;
    }
    write-host ("Already installed '{0}'" -f $_.PackageIdentifier)
}

#VsCode
$vsBackupPath = "C:\Users\{0}\AppData\Local\Temp\vsExport.json" -f $userName;
if (-not [System.IO.File]::Exists($vsBackupPath)){
    $commandOutput = & code --list-extensions
    $commandOutput.Split([Environment]::NewLine) | ConvertTo-Json | Out-File -Encoding utf8 -FilePath $vsBackupPath
}

$vsExtensionsList = Get-Content $vsBackupPath | ConvertFrom-Json
$vsExtensionsListActual = & code --list-extensions
$vsExtensionsList | Where-Object -Filter {-not $vsExtensionsListActual.Contains($_)} | ForEach-Object {
    write-host $_
    code --install-extension $_
}

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

        if ($DefaultValue -ne $null){
            $newKey["Value"] = $DefaultValue
        }

        New-Item @newKey
    } else {
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
            Path = $ItemPath
            Name = $ItemName
            Value = $ItemValue
            Type = $ItemType
        }

        New-ItemProperty @newItem
    } else {
        Write-Host "exist" + $ItemPath + "[" + $ItemType + "]"+ $ItemName + ":" + $ItemValue
    }
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


Write-Host "DONE :)"