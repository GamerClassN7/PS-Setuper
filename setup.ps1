$osName = (Get-ComputerInfo -Property OsName).OsName
if (-not [System.IO.File]::Exists("C:\Users\JonatanRek\AppData\Local\Temp\wgExport.json")){
    winget export -o "C:\Users\JonatanRek\AppData\Local\Temp\wgExport.json" >> $null
}

$wingetAppList = Get-Content "C:\Users\JonatanRek\AppData\Local\Temp\wgExport.json" | ConvertFrom-Json
$wingetAppList.Sources.Packages | ForEach-Object {
    winget list $_.PackageIdentifier >> $exist
    if ($exist -eq "No installed package found matching input criteria."){
        write-host $_.PackageIdentifier
        winget install --id="$_.PackageIdentifier"
        return;
    }
    write-host ("Already installed '{0}'" -f $_.PackageIdentifier)
}

if ($osName.startsWith("Microsoft Windows 11")){
    #Old Context Menu
    if (-not ((Get-ChildItem -path HKCU:\Software\Classes\CLSID | Where-Object -Filter {$_.Name -eq "{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"} | Measure-Object).Count -eq 0)){
        New-Item –Path "HKCU:\Software\Classes\CLSID" –Name "{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}"
        New-Item –Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" –Name "InprocServer32"
        Set-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" -Value ""
    }

    #Start Menu to left
    if (-not ((Get-ChildItem -path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" | Where-Object -Filter {$_.Name -eq "TaskbarAl"} | Measure-Object).Count -eq 0)){
        New-Item –Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" –Name "TaskbarAl" -ItemType "DWORD" -Value 0
    }
}

#Create Data Partition
#if one disk select first
# $freeDiskSpaceGb = ([math]::floor((Get-Volume -DriveLetter C).SizeRemaining /1GB) - 1)
# if ($freeDiskSpaceGb -gt 10){
#     New-Partition -DiskNumber 0 -Size 10GB -DriveLetter V | Format-Volume -FileSystem NTFS -NewFileSystemLabel "WORK 2"
# } 
