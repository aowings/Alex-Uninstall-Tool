$yes1 =0;
$yes2 =0;
$yes3 =0;
$yes4 =0;
Write-host "Hello! This tool is designed to help you completely 
remove Microsoft Office products 2010 and higher.
Please save any documents you may have open before continuing."-ForegroundColor Blue
Write-host "Would you like to to continue (Default is No)" -ForegroundColor Yellow 
    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y {Write-host "Yes"; $yes1=$true} 
       N {Write-Host "No"; $yes1=$false} 
       Default {Write-Host "Default, Skip PublishSettings"; $yes1=$false} 
     }
     if($yes1=$true){
      Write-host "Do you use Outlook?"-ForegroundColor Blue
     $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y {Write-host "Yes"; $yes2=$true} 
       N {Write-Host "No"; $yes2=$false} 
       Default {Write-Host "Default, Skip PublishSettings"; $yes2=$false} 
     }}else{
      Write-host "Quitting" -ForegroundColor Pink
      exit
     }
     if($yes2=$true){
      Write-host "Does your profile currently work?"-ForegroundColor Blue
     $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y {Write-host "Yes"; $yes3=$true} 
       N {Write-Host "No"; $yes3=$false} 
       Default {Write-Host "Default, Skip PublishSettings"; $yes3=$false} 
     } } elseif($yes2=$false){
      Write-host "do you want to start the process?"-ForegroundColor Blue
     $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y {Write-host "Yes"; $yes4=$true} 
       N {Write-Host "No"; $yes4=$false} 
       Default {Write-Host "Default, Skip PublishSettings"; $yes4=$false} 
     } }elseif($yes3=$true){
      $logFile = 'outlookprofiles.csv'

If (Test-Path 'hkcu:\Software\Microsoft\Office\15.0\Outlook\Profiles') {
	$regPath = 'hkcu:\Software\Microsoft\Office\15.0\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*'
}elseif(Test-Path 'hkcu:\Software\Microsoft\Office\16.0\Outlook\Profiles'){
	$regPath = 'hkcu:\Software\Microsoft\Office\16.0\Outlook\Profiles\*\9375CFF0413111d3B88A00104B2A6676\*'
}else{
	Write-Host 'Outlook not found'
	exit
}

Get-ItemProperty  $regPath |
	Where-Object{$_.'Account Name' -notmatch 'Outlook Address Book|Outlook Data File'} |
	Select-Object @{n= 'ComputerName';e={$env:COMPUTERNAME}},@{n= 'UserName';e={$env:USERNAME}}, 'Account Name' |
	Export-Csv $logFile 
     }
     ForEach-Object{Add-Content -Path .\outlookprofiles.csv}
     Get-Content -Path .\outlookprofiles.csv
