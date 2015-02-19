## TaxOne Auto Install v0.1##
# Robert Ostwald
# Jan 2015

##### FUNCTIONS ######

function Write-Log() {
  param(
  	[string]$text
  )
  
  $text = $text -replace("")
  $date = get-date -uformat "%a %b %d %H:%M:%S.0 %Y"
  
  $writeline = ($date + "	 " + $text)
  

  $LogFile = "C:\TaxOneInstall\TaxOneInstallation.log"
 
  out-file -filepath $LogFile -Encoding OEM -inputobject $writeline -append -noclobber
}

function Stop-Script() {
    Write-ToBox "Script stopped, please check the logfile in C:\TaxOneAutoinstall"
    exit
}

function Create-DNS () {
         
    $form_sitename = New-Object System.Windows.Forms.Form
    $tf_sitename = New-Object System.Windows.Forms.TextBox
    $hint_sitename = New-Object System.Windows.Forms.Label
	$hint_dbinst = New-Object System.Windows.Forms.Label
    $bt_sitename = New-Object System.Windows.Forms.Button
	$tf_dbinst = New-Object System.Windows.Forms.TextBox


    $font = New-Object System.Drawing.Font ("Comic Sans MS", 8)

    $tf_sitename.setBounds(10,120,200,60)
	$tf_dbinst.setBounds(10,180,200,60)

    $hint_sitename.Text = "Please enter the website's dns name and the Name of the Database Server (without any protocol or domain settings) `n`nExample: `ntaxone-test01"
    $hint_sitename.SetBounds(10,20,300,100)
    $hint_sitename.Font = $font
	
	$hint_dbinst.Text = "Name of Database Server:"
	$hint_dbinst.SetBounds(10,160,200,100)
	$hint_dbinst.Font = $font

    $bt_sitename.Text = "Ok"
    $bt_sitename.SetBounds(10,215,60,25)
    $bt_sitename.Font = $font


    $bt_sitename.add_Click({
        $dnsname = $tf_sitename.get_Text()
		$dbserv = $tf_dbinst.get_Text()
        if ($dnsname -like "" -or $dnsname -like "*.com" -or $dnsname -like "*http*" -or $dbserv -like "" -or $dbserv -like "*.DE.*") {
            Write-ToBox "Please enter a valid DNS Name and database servername"
        } 
        else {
        Write-ToBox $dnsname
        Set-Variable dnsname -value $dnsname -Scope Global
		Set-Variable dbserv -Value $dbserv -Scope global
        $form_sitename.Close()
        }
    })

    $form_sitename.Controls.Add($tf_sitename)
    $form_sitename.Controls.Add($hint_sitename)
    $form_sitename.Controls.Add($bt_sitename)
	$form_sitename.Controls.Add($tf_dbinst)
	$form_sitename.Controls.Add($hint_dbinst)
    $form_sitename.ShowDialog()
    }
	
function get-IPAdress () {
	
	$ipcon = Get-WmiObject win32_networkadapterconfiguration | ? {$_.IPEnabled}
	set-variable IPAddress -Value $ipcon.IPAddress -Scope global
}

function Write-ToBox() {
	param (
 		[string]$text
	)

[System.Windows.Forms.Application]::DoEvents()
 
$rtb_output.Lines = $rtb_output.Lines + $text
$rtb_output.SelectionStart = $rtb_output.Text.Length;$rtb_output.ScrollToCaret();
$mainForm.Refresh()

$text = ""
}

##### VARIABLES #####

## Setting folder and creating log file
mkdir "C:\TaxOneInstall" -ErrorAction SilentlyContinue
$LogFile = "C:\TaxOneInstall\TaxOneInstallation.log"

## Get Domain
Add-Type -AssemblyName System.DirectoryServices.AccountManagement 
$dom = "LDAP://" + ([ADSI]"").distinguishedName  
Write-log "Domain: $dom"

## Load Windows Forms Assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

## Global Variables
$Global:accName = $null 
$Global:pw = $null
$Global:sourcePathDB = $null
$Global:sourcePathWeb = $null
$Global:dnsname = $null
$Global:IPAddress = $null
$Global:DBServ = $null
$servername = $ENV:COMPUTERNAME

## Set IP Address Variable
get-IPAdress
$ip = $Global:IPAddress

## Set APPCMD Path
$appcmd = "C:\Windows\System32\inetsrv\appcmd.exe" 
Write-log "Path to IIS appcmd.exe $appcmd"

$tempPath = "C:\TaxOneInstall\TaxOne_DBSources"

## GUI Objects
$mainForm = New-Object System.Windows.Forms.Form

$tf_SourceDB = New-Object System.Windows.Forms.TextBox
$tf_SourceWeb = New-Object System.Windows.Forms.TextBox
$cb_Instance = New-Object System.Windows.Forms.Combobox 
$tb_database = New-Object System.Windows.Forms.Textbox
$rtb_output = New-Object System.Windows.Forms.RichTextBox

$bt_Manual = New-Object System.Windows.Forms.Button
$bt_SourceDB = New-Object System.Windows.Forms.Button
$bt_SourceWeb = New-Object System.Windows.Forms.Button
$bt_DB = New-Object System.Windows.Forms.Button
$bt_DBex = New-Object System.Windows.Forms.Button
$bt_setSVC = New-Object System.Windows.Forms.Button
$bt_webmod = New-Object System.Windows.Forms.Button
$bt_exit = New-Object System.Windows.Forms.Button

$Manuel_tooltip = New-Object System.Windows.Forms.ToolTip
$DB_toolTip = New-Object System.Windows.Forms.ToolTip
$DBex_toolTip = New-Object System.Windows.Forms.ToolTip
$webmod_toolTip = New-Object System.Windows.Forms.ToolTip
$exit_toolTip = New-Object System.Windows.Forms.ToolTip
$update_toolTip = New-Object System.Windows.Forms.ToolTip
$setSVC_toolTip = New-Object System.Windows.Forms.ToolTip
$cb_Tooltip = New-Object System.Windows.Forms.ToolTip
$tb_Tooltip = New-Object System.Windows.Forms.ToolTip

$header = New-Object System.Windows.Forms.Label
$headerFont = New-Object System.Drawing.Font ("Arial", 18, [System.Drawing.FontStyle]::Bold)
$hint = New-Object System.Windows.Forms.Label
$hint_sourceFont = New-Object System.Drawing.Font ("Comic Sans MS", 9)
$hint2 = New-Object System.Windows.Forms.Label
$hint3 = New-Object System.Windows.Forms.Label
$hint4 = New-Object System.Windows.Forms.Label
$hint5 = New-Object System.Windows.Forms.Label
$hintTFfont = New-Object System.Drawing.Font ("Comic Sans MS", 9, [System.Drawing.FontStyle]::Underline)
$hintFont = New-Object System.Drawing.Font ("Arial", 9, [System.Drawing.FontStyle]::Italic)

Write-ToBox "TaxOne AutoInstallation`nby Robert Ostwald`n`nIP: $ip `nServername: $servername `nDomain: $dom `n------------------------------------`n`nLogfile: $logFile`n"

## SQL Instance Check
if (test-path "D:\Microsoft SQL Server") {
    Write-ToBox "SQL Server Installation found. Deactivating Webservice Installation Components...`nFollow the instructions below`n`n"
	$tf_sourceWeb.Enabled = $false
	$bt_sourceWeb.Enabled = $false
	$bt_webmod.Enabled = $false
	## Add Snap-Ins for several commands for SQL Server
    Add-PSSnapin SqlServerCmdletSnapin100 -ErrorAction SilentlyContinue
    Add-PSSnapin SqlServerProviderSnapin100 -ErrorAction SilentlyContinue
    ## Checking Instances
    $instances = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances 
	foreach ($inst in $instances) {$cb_Instance.Items.Add($inst)}
    Write-Log "Installed Instances :$instances"
    Write-Log "Servername :$servername" 
	Write-ToBox "1) Please choose one of the installed Instances.`n`n2) Select the path, where the TaxOne DB Sources are located`n`n3) Define the name of the database in your chosen Instance`n"
	Write-ToBox "4) Click `"Create`" to construct a new database or click `"DB already exists`" to continue with the execution of the TaxOne DB Queries.`n"
	Write-ToBox "5) Set the Service Account, in which context the application will access the database. You need to `"Save & Close`" the prompt, when you set the credentials.`n"
} 
else {
    Write-ToBox "No SQL Server installed. Preparing for IIS Configuration...`n1) Select the path, where the web sources are located (manually or via `"Browse...`")"
	Write-ToBox "2) Please set the Instance and the database manually `n3) Set the Service Account for the Application Pool `n4) Click `"Install and Configure Application`""
	$tf_sourceDB.Enabled = $false
	$bt_sourceDB.Enabled = $false
	$bt_DB.Enabled = $false
}



####STRG V#####



## Main Window
$mainForm.Size = New-Object System.Drawing.Size @(600,650)
$mainForm.AutoSize = $true
$mainForm.TopMost = $true
$mainForm.Text = "SQL Instance Name"

	## Header for Window
	$header.setbounds(20,10,300,50)
	$header.Text = "TaxOne Installation"
	$header.Font = $headerFont

	## Hints in Window
	$hint.Text = "Please make sure, that you have administrative rights and SQL Server 2008R2 is installed on D:\."
	$hint.SetBounds(20,50,300,50)
	$hint.Font = $hintFont

	$hint2.Text = "SQL Instance"
	$hint2.SetBounds(20,150,100,50)
	$hint2.Font = $hintTFfont
	
	$hint3.Text = "Database"
	$hint3.SetBounds(20,200,100,20)
	$hint3.Font = $hintTFfont

	$hint4.Text = "Database Sources Directory"
	$hint4.SetBounds(20,50,200,20)
	$hint4.Font = $hint_sourceFont

	$hint5.Text = "Web Sources Direcotry"
	$hint5.SetBounds(20,100,200,20)
	$hint5.Font = $hint_sourceFont
	
## Source Paths
$tf_SourceDB.SetBounds(20,70,250,50)

$tf_SourceWeb.SetBounds(20,120,250,50)

$bt_SourceDB.SetBounds(280,70,100,22)
$bt_SourceDB.Text = "Browse..."

$bt_SourceWeb.SetBounds(280,120,100,22)
$bt_SourceWeb.Text = "Browse..."

## "Database" for Input
$tb_database.SetBounds(20,220,200,50)
$tb_Tooltip.ToolTipIcon = "info"
$tb_Tooltip.ToolTipTitle = “Database”
$tb_Tooltip.SetToolTip($tb_database, “The name of the database you want to create”)

## "SQL Instance" for Input
$cb_Instance.SetBounds(20,170,200,50)
$cb_Tooltip.ToolTipIcon = "info"
$cb_Tooltip.ToolTipTitle = “SQL Instance”
$cb_Tooltip.SetToolTip($cb_Instance, “Please choose one of your installed Instances”)

## Output Textbox
$rtb_output.SetBounds(20,330,450,210)
$rtb_output.ReadOnly = $true
$rtb_output.Multiline = $true
$rtb_output.Visible = $true

## "Manual" Button
$bt_Manual.SetBounds(340,60,100,40)
$bt_Manual.Text = "Manual"
$Manuel_tooltip.AutomaticDelay = 0
$Manuel_tooltip.ToolTipIcon = “warning”
$Manuel_tooltip.ToolTipTitle = “Script Guide”
$Manuel_tooltip.SetToolTip($bt_Manual, "Open a Manual of the Script”)

## "CreateDB" Button
$bt_DB.SetBounds(230,220,100,22)
$bt_DB.Text = "Create"
$DB_toolTip.AutomaticDelay = 0
$DB_toolTip.ToolTipIcon = “info”
$DB_toolTip.ToolTipTitle = “CreateDB Script”
$DB_toolTip.SetToolTip($bt_DB, “Creating a new database with the neccessary TaxOne data")
$DBFont = New-Object System.Drawing.Font ("Arial", 9, [System.Drawing.FontStyle]::Bold)
$bt_DB.Font = $DBFont

## "DB Already exists" Button
$bt_DBex.Text = "Database already exists"
$bt_DBex.SetBounds(340,220,100,22)
$bt_DBex.AutoSize = $true
$DBex_toolTip.AutomaticDelay = 0
$DBex_toolTip.ToolTipIcon = “info”
$DBex_toolTip.ToolTipTitle = “Database already exists”
$DBex_toolTip.SetToolTip($bt_DBex, “If the database is already created in the selected Instance”)

## "Set Service Account" Button
$bt_setSVC.SetBounds(20,260,100,50)
$bt_setSVC.Text = "Set Service Account"
$setSVC_toolTip.AutomaticDelay = 0
$setSVC_toolTip.ToolTipIcon = “info”
$setSVC_toolTip.ToolTipTitle = “Service Account Configuration”
$setSVC_toolTip.SetToolTip($bt_setSVC, “Set the Service Account giving him all neccessary permissions.")

## "Install and Configure Application"
$bt_webmod.SetBounds(140,260,150,50)
$bt_webmod.Text = "Install and Configure Application"
$webmod_toolTip.SetToolTip($bt_webmod, "You need to specify a Service Account in which context the application pool will run.")
$webmod_toolTip.AutomaticDelay = 0

## "Exit" Button
$bt_exit.Text = "Exit"
$bt_exit.SetBounds(20,560,100,23)
$exit_toolTip.AutomaticDelay = 0
$exit_toolTip.ToolTipIcon = “error”
$exit_toolTip.ToolTipTitle = “Exit the Setup”
$exit_toolTip.SetToolTip($bt_exit, “Nope. No way I'm using that.")

# Add "Browse..." Click Events
$bt_SourceDB.add_Click({

    $fbd_selectSourcePathDB = New-Object -com Shell.Application
    $folder = $fbd_selectSourcePathDB.BrowseForFolder(0, "Select the location of SQL scripts", 0, "Computer")
    $sourcePathDB = $folder.Self.Path
    Set-Variable sourcePathDB -Value $sourcePathDB -Scope global
    Write-Log "Sources for DB: $sourcePathDB"
    $tf_SourceDB.AppendText($sourcePathDB)
    
})

$bt_SourceWeb.add_Click({
    $fbd_selectSourcePathW = New-Object -com Shell.Application
    $folder = $fbd_selectSourcePathW.BrowseForFolder(0, "Select the location of web sources", 0, "Computer")
    $sourcePathWeb = $folder.Self.Path

    Set-Variable sourcePathWeb -Value $sourcePathWeb -Scope global
    Write-Log "Sources for Web: $sourcePathWeb"
    $tf_SourceWeb.AppendText($sourcePathWeb)
    
})
 
# Add "Create" Click Event
$bt_DB.add_Click({
	
    ## Path
    $sourcePathDB = $tf_SourceDB.get_Text()
    
	## Get Instance Name for Script	
	$instName = $cb_Instance.get_Text()
	
	## Get DB Name and Create File
	$newDB = $tb_database.get_Text()

    ## Check path, instance and database name
	if ($newDB -like "" -or $instName -like "" -or $sourcePathDB -like "") {
		Write-ToBox "You need to choose an Instance, set a Name for the database and choose the path to the SQL queries."
	}
    ## When everthing need is set, the script will be created
	else {
		$newDBfile = "CreateDB.sql"
        $newDBlog = "$newDB" + '_log'
            

		## Script for creating database
		$y = Test-Path $tempPath 
		if ($y -like $false) {
			mkdir $tempPath
            Write-Log "Creating $tempPath" 
		}
			
		Out-File -FilePath "$tempPath\$newDBfile" -Force
		Set-ItemProperty -Path "$tempPath\$newDBfile" -Name IsReadOnly -Value $false
		Write-Log "$newDBFile has been created at $tempPath" 
        sleep -Seconds 3
	    Add-Content -Encoding Unknown -Path "$tempPath\$newDBfile" -Value "CREATE DATABASE [$newDB] ON  PRIMARY 
( NAME = N'$newDB', FILENAME = N'D:\Microsoft SQL Server\MSSQL10_50.$instName\MSSQL\DATA\$newDB.mdf' , SIZE = 5120KB , FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'$newDBlog', FILENAME = N'E:\Microsoft SQL Server\MSSQL10_50.$instName\MSSQL\Data\$newDBlog.ldf' , SIZE = 1024KB , FILEGROWTH = 10%)
GO
ALTER DATABASE [$newDB] SET COMPATIBILITY_LEVEL = 100
GO
ALTER DATABASE [$newDB] SET ANSI_NULL_DEFAULT OFF 
GO
ALTER DATABASE [$newDB] SET ANSI_NULLS OFF 
GO
ALTER DATABASE [$newDB] SET ANSI_PADDING OFF 
GO
ALTER DATABASE [$newDB] SET ANSI_WARNINGS OFF 
GO
ALTER DATABASE [$newDB] SET ARITHABORT OFF 
GO
ALTER DATABASE [$newDB] SET AUTO_CLOSE OFF 
GO
ALTER DATABASE [$newDB] SET AUTO_CREATE_STATISTICS ON 
GO
ALTER DATABASE [$newDB] SET AUTO_SHRINK OFF 
GO
ALTER DATABASE [$newDB] SET AUTO_UPDATE_STATISTICS ON 
GO
ALTER DATABASE [$newDB] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO
ALTER DATABASE [$newDB] SET CURSOR_DEFAULT  GLOBAL 
GO
ALTER DATABASE [$newDB] SET CONCAT_NULL_YIELDS_NULL OFF 
GO
ALTER DATABASE [$newDB] SET NUMERIC_ROUNDABORT OFF 
GO
ALTER DATABASE [$newDB] SET QUOTED_IDENTIFIER OFF 
GO
ALTER DATABASE [$newDB] SET RECURSIVE_TRIGGERS OFF 
GO
ALTER DATABASE [$newDB] SET  DISABLE_BROKER 
GO
ALTER DATABASE [$newDB] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO
ALTER DATABASE [$newDB] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO
ALTER DATABASE [$newDB] SET PARAMETERIZATION SIMPLE 
GO
ALTER DATABASE [$newDB] SET  READ_WRITE 
GO
ALTER DATABASE [$newDB] SET RECOVERY FULL 
GO
ALTER DATABASE [$newDB] SET  MULTI_USER 
GO
ALTER DATABASE [$newDB] SET PAGE_VERIFY CHECKSUM  
GO
USE [$newDB]
GO
IF NOT EXISTS (SELECT name FROM sys.filegroups WHERE is_default=1 AND name = N'PRIMARY') ALTER DATABASE [$newDB] MODIFY FILEGROUP [PRIMARY] DEFAULT
GO
"	

        sleep -Seconds 3
        Write-ToBox -ForegroundColor DarkGreen "SQL Script successfully created in $tempPath"
        Write-Log "Script for Creation of Database: $newDBfile at $tempPath"
	
        ## Creating Database
        Write-ToBox "Executing SQLCMD Command to Create Database"
        try {
            Invoke-sqlcmd -serverInstance $servername\$instName -InputFile $tempPath\$newDBfile -queryTimeout ([int]::MaxValue)
        } 
        catch  {
            Write-Log "$_"
            Write-ToBox "Failed to execute $newDBfile. Please check the logfile for more information."
            Stop-Script
        }
        Write-ToBox "Creation of Database successful. Continuing with Querys for TaxOne"
        sleep -Seconds 3
    
        ## Executing TaxOne Querys
        $queries = Get-ChildItem $sourcePathDB -Filter "*.sql"
        Write-Log "Path: $sourcePathDB"
        for ($i=0; $i -lt $queries.count; $i++) {
            $query = $queries[$i].Name
            try { 
                Write-ToBox "Executing $query"
                Invoke-sqlcmd -serverInstance $servername\$instName -Database $newDB -inputFile $sourcePathDB\$query -queryTimeout ([int]::MaxValue)
                Write-log "Invoke-sqlcmd -serverInstance $servername\$instName -database $newDB -inputFile $sourcePathDB\$query"
                sleep -Seconds 3
            }
            catch {
                Write-Log "$_"
                Write-ToBox "Failed to execute $query. Check the Logfile for more information"
                Stop-Script
            }
        }

        Write-Log "SQL Installation done" 
        Write-ToBox "SQL Installation done. Please continue by clicking `"Set SQL Service Account`""

    } #end of Else

})	

# Add "DB already exists" Click Event
$bt_DBex.add_Click({
    if (($tb_database.get_Text()) -like "") {
		Write-ToBox "Please define a database name"
	}
	else {
		## Executing TaxOne Querys
        $queries = Get-ChildItem $sourcePathDB -Filter "*.sql"
        Write-Log "Path: $sourcePathDB"
        for ($i=0; $i -lt $queries.count; $i++) {
            $query = $queries[$i].Name
            try { 
                Write-ToBox "Executing $query"
                Invoke-sqlcmd -serverInstance $servername\$instName -Database $newDB -inputFile $sourcePathDB\$query -queryTimeout ([int]::MaxValue)
                Write-log "Invoke-sqlcmd -serverInstance $servername\$instName -database $newDB -inputFile $sourcePathDB\$query"
                sleep -Seconds 3
            }
            catch {
                Write-Log "$_"
                Write-ToBox "Failed to execute $query. Check the Logfile for more information"
                Stop-Script
            }
        }

        Write-Log "SQL Installation done" 
        Write-ToBox "SQL Installation done. Please continue by clicking `"Set SQL Service Account`""
	}

    $bt_DB.Enabled = $false
})

# Add "Set Service Account" Click Event
$bt_setSVC.add_Click({

    ## Get Instance Name for Script	
	$instName = $cb_Instance.get_Text()
	
	## Get DB Name
	$newDB = $tb_database.get_Text()	
    
    ## Menue for Service Account
    $setSVC_Form = New-Object System.Windows.Forms.Form
    $setSVC_Form.Size = New-Object System.Drawing.Size @(400,250)
    $setSVC_Form.Text = "Service Account"
    $setSVC_Form.TopMost = $true

    ## Create Hint for Accoutname
    $svc_hint1 = New-Object System.Windows.Forms.Label
    $svc_hint1.Text = "Accountname"
    $svc_hint1.SetBounds(20,50,250,20)
    
    ## Create Hint for Password
    $svc_hint2 = New-Object System.Windows.Forms.Label
    $svc_hint2.Text = "Pasword"
    $svc_hint2.SetBounds(20,100,250,20)

    ## Create Account Textbox
    $tb_acc = New-Object System.Windows.Forms.Textbox
    $tb_acc.SetBounds(20,70,200,50)

    ## Create Password Textbox
    $tb_pw = New-Object System.Windows.Forms.Textbox
    $tb_pw.SetBounds(20,120,200,50)
    $tb_pw.PasswordChar = "*"
    
    ## Create "Check" Button
    $bt_svcCheck = New-Object System.Windows.Forms.Button
    $bt_svcCheck.Text = "Check"
    $bt_svcCheck.SetBounds(20,160,100,30)

    ## Create "Save & Close" Button
    $bt_svcSave = New-Object System.Windows.Forms.Button
    $bt_svcSave.Text = "Save and Close"
    $bt_svcSave.SetBounds(130,160,100,30)
    $bt_svcSave.Enabled = $false

    ## Create "Close" Button
    $bt_svcClose = New-Object System.Windows.Forms.Button
    $bt_svcClose.text = "Close"
    $bt_svcClose.SetBounds(240,160,100,30)

    ## Add "Check" Click Event
    $bt_svcCheck.add_Click({

        ## Checking Account to Domain
        $accName = $tb_acc.get_Text()
        $pw = $tb_pw.get_Text()
        $creds = New-Object System.DirectoryServices.DirectoryEntry($dom,$accName,$pw)
        if ($creds.name -eq $null) {
            [System.Windows.Forms.MessageBox]::Show("Authentification failed.","Check",0)
	    }
	    else {
	        [System.Windows.Forms.MessageBox]::Show("Authentification successful.","Check",0)
            $bt_svcSave.Enabled = $true
            $bt_webmod.Enabled = $true
            ## Setting global variable $accName
            Set-Variable accName -Value "$accName" -Scope Global
            Set-Variable pw -Value "$pw" -Scope global 
            
	    }

    })

## Add "Save & Close" Event
    $bt_svcSave.add_Click({

        $accName = $tb_acc.get_Text()
        $pw = $tb_pw.get_Text()
        
        ## Creating SQL Command to give permissions to Service Account
        out-file -FilePath "$tempPath\$accName.sql"
	    Add-Content -Encoding Unknown -Path "$tempPath\$accName.sql" -Value "USE [master]
GO
CREATE LOGIN [DE\$accname] FROM WINDOWS WITH DEFAULT_DATABASE=[master]
GO
USE [$newDB]
GO
CREATE USER [de\$accName] FOR LOGIN [de\$accName]
GO
USE [$newDB]
GO
EXEC sp_addrolemember N'taxportal_role', N'de\$accName'
GO"
        Invoke-sqlcmd -serverInstance $servername\$instname -Database $newDB -inputFile "$tempPath\$accName.sql" -querytimeout ([int]::MaxValue)
        Write-Log "Invoke-sqlcmd -serverInstance $servername\$instname -Database $newDB -inputFile `"$tempPath\$accName.sql`" -querytimeout ([int]::MaxValue)"
        Write-Log "$accName successfully allowed on $newDB"
        Write-ToBox "SQL permissions granted."

        $setSVC_Form.Close()
    })

    ## Add "Close" Event
    $bt_svcClose.add_Click({
        $setSVC_Form.Close()
    })

    $setSVC_Form.Controls.Add($svc_hint1)
    $setSVC_Form.Controls.Add($svc_hint2)
    $setSVC_Form.Controls.Add($tb_acc)
    $setSVC_Form.Controls.Add($tb_pw)
    $setSVC_Form.Controls.Add($bt_svcCheck)
    $setSVC_Form.Controls.Add($bt_svcSave)
    $setSVC_Form.Controls.Add($bt_svcClose)
    $setSVC_Form.ShowDialog()    
})

# Add "Install and Configure Application" Click Event
$bt_webmod.add_Click({
    
    $accName = $global:accName
    $pw = $global:pw
	Write-ToBox "Specified account: $accname passsword: $pw"
    $sourcePathWeb = $tf_SourceWeb.get_Text()
    $instance = $cb_instance.get_Text()
    $newDB = $tb_database.get_Text()
	
	Get-IPAdress
    create-dns

    $sitename = $Global:dnsname
	$dbserv = $Global:dbserv
	
    $destPathWeb = "D:\WebServer\$sitename\GeneralWebsite"

    if ($instance -eq "" -or $instance -eq $null -or $sourcePathWeb -eq "" -or $sourcePathWeb -eq $null -or $newDB -eq "" -or $newDB -eq $null -or $accName -eq "" -or $accName -eq $null) {
        Write-ToBox "Please make sure, that the instance, the database, the source directory and the service Account have been set."
    } 
    else {
        if (Test-Path $destPathWeb) {}
        else {
            mkdir $destPathWeb -Force
        }
        copy-item -Path $sourcePathWeb\* -Destination $destPathWeb -Recurse
        
        Write-ToBox "Start copying source files to $destpathweb"
		Write-Log "Connection string will be set to <property name=`"connection.connection_string`">
          Server=$servername\$instance;initial catalog=$newDB;Integrated Security=SSPI;Min Pool Size=1"
		$modules = Get-ChildItem -Path "$destPathWeb"

        for ($i=0; $i -le $modules.count; $i++) {
            $modul = $modules[$i].Name
            Remove-Item "$destPathWeb\$modul\" -Include "*sample*" -Recurse 
			mkdir "$destPathWeb\$modul\Log"
            Write-ToBox "Replacing $modul hibernate.cfg.xml"
       	    Out-File "$destPathWeb\$modul\hibernate.cfg.xml" -Force -Encoding utf8 -ErrorAction SilentlyContinue
            sleep -Seconds 2
		    Set-ItemProperty -Path "$destPathWeb\$modul\hibernate.cfg.xml" -Name IsReadOnly -Value $false
		    Write-Log "$destPathWeb\$modul\hibernate.cfg.xml has been created"
            sleep -Seconds 3
            Add-Content -Encoding UTF8 -Path "$destPathWeb\$modul\hibernate.cfg.xml" -Value "<?xml version='1.0' encoding='utf-8'?>
    <hibernate-configuration xmlns=`"urn:nhibernate-configuration-2.2`">
      <session-factory>
        <property name=`"dialect`">NHibernate.Dialect.MsSql2005Dialect</property>
        <property name=`"connection.provider`">NHibernate.Connection.DriverConnectionProvider</property>
        <property name=`"connection.driver_class`">NHibernate.Driver.SqlClientDriver</property>
        <property name=`"connection.connection_string`">
          Server=$dbserv\$instance;initial catalog=$newDB;Integrated Security=SSPI;Min Pool Size=1
        </property>
        <property name=`"connection.isolation`">ReadCommitted</property>
      </session-factory>
    </hibernate-configuration>"

        }	
	
        $x = $appcmd.ToString()

        ## Create Application Pool
        Write-ToBox "Starting IIS Configuration. View Log for further information"
        start "$x" -ArgumentList "add apppool -name:`"$sitename`""
        Write-Log "Application Pool created: $sitename"
        sleep -Seconds 5

        ## Change ApplicationPool in .NET 4.0 and Classic Pipeline Mode
        start "$x" -ArgumentList "set apppool `"$sitename`" /managedPipelineMode:Classic"
        sleep -Seconds 5
        start "$x" -ArgumentList "set apppool `"$sitename`" /managedRuntimeVersion:v4.0"
        sleep -Seconds 5
		## Set localservice as Service Account for AppPool
        start "$x" -ArgumentList "set config -section:system.applicationHost/applicationPools -[name=`'$sitename`'].processModel.identityType:`"LocalService`" -commit:apphost"
        sleep -Seconds 2

        
        ## Bindings
        $dnscheck = [System.Windows.forms.MessageBox]::Show("DNS Entry already set?", "DNS Check", 4) 
        Write-Log "DNS entry answered with $dnscheck"
        ## If DNS Entry is set, the bindings will be specified with the dns entry, otherwise, the site will be bound over *:80
        if ($dnscheck -eq "Yes") {
            start "$x" -ArgumentList "add site -name:`"$sitename`" -bindings:`"HTTP/$ip`:80`:$sitename`" -physicalpath:`"$destPathWeb`"" 
            Write-Log "Creating website '$sitename' with binding 'HTTP/$ipaddress:80:$sitename' and at $destPathWeb"
            sleep -Seconds 2
            start "$x" -ArgumentList "set site `"$sitename`" /+bindings.[protocol='http',bindingInformation=`'$ip`:80`:$sitename.de.kworld.kpmg.com']"
            Write-Log "Set Bindings: http*:80:$sitename.de.kworld.kpmg.com']"
            sleep -Seconds 2
			
			## Create index.html
			Out-File "$destPathWeb\index.html" -Force -Encoding utf8 -ErrorAction SilentlyContinue
			Write-Log "Creating index.html @ $destPathWeb"
			sleep -Seconds 2
			Add-Content -Encoding utf8 -Path "$destPathWeb\index.html" -Value "<html>
			<head>
				<title>TaxOne</title>
			</head>
			<body>
		<SCRIPT>
			location.replace(`'http://$sitename.de.kworld.kpmg.com/taxportal`');
		</SCRIPT>
			</body>
		</html>"
		
         }
         else {
            start "$x" -ArgumentList "add site -name:`"$sitename`" -physicalpath:`"$destPathWeb`""
            sleep -Seconds 2
            start "$x" -ArgumentList "set site `"$sitename`" /+bindings.[protocol='http',bindingInformation='$ip`:80`:']" 
            sleep -seconds 2
         }
		 

        ## Set application pool identity
		start "$x" -ArgumentList "stop apppool -name:`"$sitename`""
		sleep -Seconds 3
 		start "$x" -argumentlist "set config /section`:applicationPools /[name=`'$sitename`'].processModel.identityType`:SpecificUser /[name=`'$sitename`'].processModel.userName`:de\$accname /[name=`'$sitename`'].processModel.password`:$pw"
		sleep -Seconds 4
		start "$x" -ArgumentList "start apppool -name:`"$sitename`""

        ## Set IIS log directory
        mkdir "E:\LogFiles\$instance" -ErrorAction SilentlyContinue
        start "$x" -ArgumentList "set config -section:system.applicationHost/sites -[name=`'$sitename`'].logfile.directory:`"E:\LogFiles\$sitename`""
		sleep -Seconds 2
				
		## Set website to use the created AppPool
		start "$x" -ArgumentList "set config -section:system.applicationHost/sites -[name=`'$sitename`'].applicationDefaults.applicationPool:`"$sitename`"  -commit:apphost"
        
        ## Convert virtual directories to applications
        Write-ToBox "Converting virtual directories to applications..."
        for ($i=0; $i -le $modules.count; $i++) {
            $modul = $modules[$i].Name
            start "$x" -ArgumentList "add app /site.name:`"$sitename`" /path:/`"$modul`" /physicalPath:`"$destPathweb\$modul`""
            Write-Log "Converting $modul to application"
            sleep -Seconds 2
        }

        Write-ToBox "Done"

        ## Set TaxPortal Default Document
        start "$x" -ArgumentList "set config `"$sitename/TaxPortal`" /section:defaultDocument `"/+files.[@start,value='TaxPortal.aspx']`" /commit:`"$sitename/TaxPortal`""
        Write-Log "Setting TaxPortal Default Document"
    }
        
})

# Add "Exit" Click Event
$bt_exit.add_Click({
	[System.Windows.Forms.MessageBox]::Show("That button has no function at all. It is just there. Deal with it.", "Nope.", 0, [System.Windows.Forms.MessageBoxIcon]::Error)
})

## Add Objects to Main Window
$mainForm.Controls.Add($cb_Instance)
$mainForm.Controls.Add($tb_database)
$mainForm.Controls.Add($bt_DBex)
#$mainForm.Controls.Add($hint)
$mainForm.Controls.Add($hint2)
$mainForm.Controls.Add($hint3)
$mainForm.Controls.Add($hint4)
$mainForm.Controls.Add($hint5)
$mainForm.Controls.Add($header)
$mainForm.Controls.Add($bt_DB)
$mainForm.Controls.Add($bt_SourceDB)
$mainForm.Controls.Add($bt_SourceWeb)
$mainForm.Controls.Add($bt_setSVC)
$mainForm.Controls.Add($bt_webmod)
$mainForm.Controls.Add($bt_exit)
$mainForm.Controls.Add($tf_SourceDB)
$mainForm.Controls.Add($tf_SourceWeb)
$mainForm.Controls.Add($rtb_output)
#$mainForm.Controls.Add($bt_Manual)
$mainForm.ShowDialog()
