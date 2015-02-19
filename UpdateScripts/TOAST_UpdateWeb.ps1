## Updating TaxOne via Script

# Load Windows Forms Assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

## Set APPCMD Path
$appcmd = "C:\Windows\System32\inetsrv\appcmd.exe"
$x = $appcmd.ToString()

## Set IE path
$iexplore = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"

## Variables

Write-Host "Please ensure, that ur TaxOne root is installed on D:\WebServer\INSTANCENAME\GeneralWebsite"


$updatepath = "\\defr2web94\D$\Lieferung\TaxOne\TaxPortal_Root" 
$cont1 = [System.Windows.Forms.MessageBox]::Show("Please make sure that the new files are located in $updatepath. By pressing `"OK`", the site belonging to the entered instance will be stopped.", "Check", 0)

while ($true) {

$instance = Read-Host -Prompt "Instance" 

$installdir = "D:\WebServer\$instance\GeneralWebSite" 

$date = Get-Date -Format yyyy-MM-dd
$bpath = "D:\WebServer\$instance\Backup\" + "$date" 

## Stop website
start "$x" -ArgumentList "stop site /site.name:$instance"
sleep -Seconds 3

## Backup old installation
Write-Host "Backing up $installdir to $bpath"

if (Test-Path $installdir) {
    Copy-Item -path $installdir -Recurse -Destination $bpath 
}
else {
    Write-Host "$installdir does not exist"
    exit
}

## Deleting old Installation
Write-Host "Removing old installation files except hibernate.cfg.xml, index.html and Logs directory for each module"
Remove-Item -path "$installdir\*\*" -Recurse -exclude hibernate.cfg.xml,Logs,index.html
sleep -Seconds 3

## Grab new file
Write-Host "Collecting new data on $updatepath"
Copy-Item -Path "$updatepath\*" -Destination $installdir -Recurse -ErrorAction SilentlyContinue -force

Write-host "Update done. Check the IIS configuration and restart it."

## Set TaxPortal Default Document
Write-Host "Setting TaxPortal Default Document"
start "$x" -ArgumentList "set config `"$instance/TaxPortal`" /section:defaultDocument `"/+files.[@start,value='TaxPortal.aspx']`" /commit:`"$instance/TaxPortal`""
sleep -Seconds 3

## Authentication Mode
Write-Host "Setting Authentification to Windows Authentification"
start "$x" -ArgumentList "unlock config /section:windowsAuthentication"
sleep -Seconds 3
start "$x" -ArgumentList "set config `"$instance`" /section:windowsAuthentication /enabled:true /commit:apphost"
sleep -Seconds 3

## Restart site
Write-Host "The website will be started now."
start "$x" -ArgumentList "start site /site.name:$instance"
sleep -Seconds 3

Write-Host "Opening Internet Explorer to test availablility..."
start $iexplore -ArgumentList "$instance.de.kworld.kpmg.com"

Write-Host "If the databases need to be updated, please execute the TaxPortalUpdateDB Script on the database server."
Write-Host "Webservice files update finished."

}

## End

