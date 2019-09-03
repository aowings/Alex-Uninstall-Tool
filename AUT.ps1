$yes1 = 0;
$yes2 = 0;
$yes3 = 0;
$yes4 = 0;
$UserKeyDel = 0;
Write-host "Hello! This tool is designed to help you completely 
remove Microsoft Office products 2010 and higher.
Please save any documents you may have open before continuing."-ForegroundColor Blue
Write-host "Would you like to to continue (Default is No)" -ForegroundColor Yellow 
$Readhost = Read-Host " ( y / n ) " 
Switch ($ReadHost) { 
    Y { Write-host "Yes"; $yes1 = $true } 
    N { Write-Host "No"; $yes1 = $false } 
    Default { Write-Host "Default, Skip PublishSettings"; $yes1 = $false } 
}
if ($yes1 = $true) {
    Write-host "Do you use Outlook?"-ForegroundColor Blue
    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) { 
        Y { Write-host "Yes"; $yes2 = $true } 
        N { Write-Host "No"; $yes2 = $false } 
        Default { Write- "Default, Skip PublishSettings"; $yes2 = $false } 
    }
}
elseif ($yes1 = $false) {
    Write-host "Quitting" -ForegroundColor Pink
    exit
}
if ($yes2 = $true) {
    Write-host "Does your profile currently work?"-ForegroundColor Blue
    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) { 
        Y { Write-host "Yes"; $yes3 = $true } 
        N { Write-Host "No"; $yes3 = $false } 
        Default { Write-Host "Default, Skip PublishSettings"; $yes3 = $false } 
    } 
}
elseif ($yes2 = $false) {
    Write-host "do you want to start the process?"-ForegroundColor Blue
    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) { 
        Y { Write-host "Yes"; $yes4 = $true } 
        N { Write-Host "No"; $yes4 = $false } 
        Default { Write-Host "Default, Skip PublishSettings"; $yes4 = $false } 
    } 
}
elseif ($yes3 = $true) {
    $logFile = 'outlookprofiles.csv'

    If (Test-Path 'hkcu:\Software\Microsoft\Office\15.0\Outlook\Profiles') {
        $regPath = 'hkcu:\Software\Microsoft\Office\15.0\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*'
    }
    elseif (Test-Path 'hkcu:\Software\Microsoft\Office\16.0\Outlook\Profiles') {
        $regPath = 'hkcu:\Software\Microsoft\Office\16.0\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*'
    }
    else {
        Write-Host 'Outlook not found'
        exit
    }

    Get-ItemProperty  $regPath |
    Where-Object { $_.'Account Name' -notmatch 'Outlook Address Book|Outlook Data File' } |
    Select-Object @{n = 'ComputerName'; e = { $env:COMPUTERNAME } }, @{n = 'UserName'; e = { $env:USERNAME } }, 'Account Name' |
    Export-Csv $logFile 
}
ForEach-Object { Add-Content -Path .\outlookprofiles.csv }
Get-Content -Path .\outlookprofiles.csv
if ($yes3 = $false){
//Automated recovery of outlook profile.
}
if ($yes4 = $true) {
    Write-host "We are clearing your current installation of office"
    Get-ChildItem -Path C:\windows\Temp -Include *.* -File -Recurse | foreach-Object { $_.Delete() }
    Get-ChildItem -Path C:\Users\$env:USERNAME\AppData\Local\Temp -Include *.* -File -Recurse | foreach-Object { $_.Delete() }
    Get-ChildItem -Path C:\Program Files (x86)\Microsoft Office -Include *.* -File -Recurse | foreach-Object { $_.Delete() }
    Get-ChildItem -Path C:\Program Files (x86)\Common Files\Microsoft Shared -Include *.* -File -Recurse | OFFICE16 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files (x86)\Common Files\Microsoft Shared -Include *.* -File -Recurse | OFFICE11 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files (x86)\Common Files\Microsoft Shared -Include *.* -File -Recurse | OFFICE12 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files (x86)\Common Files\Microsoft Shared -Include *.* -File -Recurse | OFFICE13 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files (x86)\Common Files\Microsoft Shared -Include *.* -File -Recurse | OFFICE14 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files (x86)\Common Files\Microsoft Shared -Include *.* -File -Recurse | OFFICE15 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files (x86)\Common Files\Microsoft Shared -Include *.* -File -Recurse | ClickToRun { $_.Delete() }
    Get-ChildItem -Path C:\Program Files\Common Files\microsoft shared -Include *.* -File -Recurse | ClickToRun { $_.Delete() }
    Get-ChildItem -Path C:\Program Files\Microsoft Office 15 -Include *.* -File -Recurse | foreach-Object { $_.Delete() }
    Get-ChildItem -Path C:\Program Files\Microsoft Office -Include *.* -File -Recurse | foreach-Object { $_.Delete() }
    Get-ChildItem -Path C:\Program Files\Common Files\microsoft shared -Include *.* -File -Recurse | OFFICE16 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files\Common Files\microsoft shared -Include *.* -File -Recurse | OFFICE11 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files\Common Files\microsoft shared -Include *.* -File -Recurse | OFFICE12 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files\Common Files\microsoft shared -Include *.* -File -Recurse | OFFICE13 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files\Common Files\microsoft shared -Include *.* -File -Recurse | OFFICE14 { $_.Delete() }
    Get-ChildItem -Path C:\Program Files\Common Files\microsoft shared -Include *.* -File -Recurse | OFFICE15 { $_.Delete() }
    Get-ChildItem -Path hkcu:\HKEY_CURRENT_USER\Software\Microsoft\Office { $_.Delete() }
    Get-ChildItem -Path hkcu:\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office { $_.Delete() }
    Get-ChildItem -Path hkcu:\HKEY_USERS\.DEFAULT\Software\Microsoft\Office { $_.Delete() }
    Get-ChildItem -Path hkcu:\HKEY_USERS\S-1-5-18\Software\Microsoft\Office { $_.Delete() }
    Get-ChildItem -Path hkcu:\HKEY_USERS\S-1-5-19\Software\Microsoft\Office { $_.Delete() }
    Get-ChildItem -Path hkcu:\HKEY_USERS\S-1-5-20\Software\Microsoft\Office { $_.Delete() }
    Get-ChildItem -Path hkcu:\HKEY_USERS\S-1-5-20\Software\Microsoft\Office { $_.Delete() }
    $UserKeyDel = Get-ChildItem Registry::HKEY_USERS\ | Where-Object { $_.Name -match '\Software\Microsoft\Office' }
    foreach ($UserKeyDel in $UserKeyDel) {
        \Software\Microsoft\Office { $_.Delete() }
    }
    else {
        # doesn't exist
    } else($yes4 = $false){
      Write-Host "Quitting"
      exit
    }
}
