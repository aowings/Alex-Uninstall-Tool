Out-String("Hello! This tool is designed to help you completely 
remove Microsoft Office products 2010 and higher.
Please save any documents you may have open before continuing.")
Write-host "Would you like to to continue (Default is No)" -ForegroundColor Yellow 
    $Readhost = Read-Host " ( y / n ) " 
    Switch ($ReadHost) 
     { 
       Y {Write-host "Yes, Download PublishSettings"; $yes=$true} 
       N {Write-Host "No, Skip PublishSettings"; $yes=$false} 
       Default {Write-Host "Default, Skip PublishSettings"; $yes=$false} 
     } 