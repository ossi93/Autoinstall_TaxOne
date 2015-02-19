## TaxOne Auto Install ##
#
##
# Execution Policy
#Set-ExecutionPolicy bypass

mkdir "C:\TaxOneInstall" -ErrorAction SilentlyContinue

## Get Domain
Add-Type -AssemblyName System.DirectoryServices.AccountManagement 
$dom = "LDAP://" + ([ADSI]"").distinguishedName

## Function Write-Log
#  Write-Log "Text" creates a new line in the log file with the date and the given text

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
## Function End

# Load Windows Forms Assembly
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

## SQL Instance Check
$instances = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances 
$servername = $ENV:COMPUTERNAME

## Global Variables
$Global:accName = $null
$Global:pw = $null

## Set APPCMD Path
$appcmd = "C:\Windows\System32\inetsrv\appcmd.exe"

## Source Path
$sourcePathDB = "\\DEFR4-RO01\appsrv$\TaxCompliance_Steuerportal_2.1417.0.101\DB"
$sourcePathWeb = "\\DEFR4-RO01\appsrv$\TaxCompliance_Steuerportal_2.1417.0.101\TaxPortal_Root"
$destPathDB = "C:\TaxOneInstall\TaxOne_DBSources"
$LogFile = "C:\TaxOneInstall\TaxOneInstallation.log"
Write-Log "DB Sources = $sourcePathDB"
Write-Log "Web Sources = $sourcePathWeb"
Write-Log "Destination = C:\TaxOneInstall\"

## Create Main Window
$mainForm = New-Object "System.Windows.Forms.Form"
$mainForm.Size = New-Object System.Drawing.Size @(500,350)
$mainForm.AutoSize = $true
$mainForm.TopMost = $true
$mainForm.Text = "SQL Instance Name"

	## Create Header for Window
	$header = New-Object System.Windows.Forms.Label
	$header.setbounds(20,10,300,50)
	$header.Text = "TaxOne Installation"
	$headerFont = New-Object System.Drawing.Font ("Arial", 18, [System.Drawing.FontStyle]::Bold)
	$header.Font = $headerFont

	## Create Hints in Window
	$hint = New-Object System.Windows.Forms.Label
	$hint.Text = "Please make sure, that you have administrative rights and SQL Server 2012 is installed on D:\."
	$hint.SetBounds(20,50,300,50)
	$hintFont = New-Object System.Drawing.Font ("Arial", 9, [System.Drawing.FontStyle]::Italic)
	$hint.Font = $hintFont

	$hint2 = New-Object System.Windows.Forms.Label
	$hint2.Text = "SQL Instance"
	$hint2.SetBounds(20,100,100,50)
	$hint2Font = New-Object System.Drawing.Font ("Arial", 9, [System.Drawing.FontStyle]::Underline)
	$hint2.Font = $hint2Font
	
	$hint3 = New-Object System.Windows.Forms.Label
	$hint3.Text = "Database"
	$hint3.SetBounds(20,150,100,20)
	$hint3Font = New-Object System.Drawing.Font ("Arial", 9, [System.Drawing.FontStyle]::Underline)
	$hint3.Font = $hint3Font
	
## Create "Database" for Input
$tb_database = New-Object System.Windows.Forms.Textbox
$tb_database.SetBounds(20,170,200,50)
$tb_Tooltip = New-Object System.Windows.Forms.ToolTip
$tb_Tooltip.ToolTipIcon = "info"
$tb_Tooltip.ToolTipTitle = “Database”
$tb_Tooltip.SetToolTip($tb_database, “The name of the database you want to create”)

## Create "SQL Instance" for Input
$cb_Instance = New-Object "System.Windows.Forms.Combobox" 
$cb_Instance.SetBounds(20,120,200,50)
$cb_Instance.Items.Add("$instances")
$cb_Tooltip = New-Object System.Windows.Forms.ToolTip
$cb_Tooltip.ToolTipIcon = "info"
$cb_Tooltip.ToolTipTitle = “SQL Instance”
$cb_Tooltip.SetToolTip($cb_Instance, “Please choose one of your installed Instances”)

## Create "Manual Button" 
$bt_Manual = New-Object System.Windows.Forms.Button
$bt_Manual.SetBounds(340,10,100,40)
$bt_Manual.Text = "Manual"
$Manuel_tooltip = New-Object System.Windows.Forms.ToolTip
$Manuel_tooltip.AutomaticDelay = 0
$Manuel_tooltip.ToolTipIcon = “warning”
$Manuel_tooltip.ToolTipTitle = “Script Guide”
$Manuel_tooltip.SetToolTip($bt_Manual, "Open a Manual of the Script”)

## Create "Not installed yet" Button
$bt_niy = New-Object "System.Windows.Forms.Button"
$bt_niy.Text = "Not installed yet"
$bt_niy.SetBounds(230,120,100,22)
$niy_toolTip = New-Object System.Windows.Forms.ToolTip
$niy_toolTip.AutomaticDelay = 0
$niy_toolTip.ToolTipIcon = “warning”
$niy_toolTip.ToolTipTitle = “SQL Installation Setup”
$niy_toolTip.SetToolTip($bt_niy, “When the Instance doesnt exist, the SQL2008R2 Installation will start”)

## Create "CreateDB" Button
$bt_DB = New-Object "System.Windows.Forms.Button"
$bt_DB.SetBounds(230,170,100,22)
$bt_DB.Text = "Create"
$DB_toolTip = New-Object System.Windows.Forms.ToolTip
$DB_toolTip.AutomaticDelay = 0
$DB_toolTip.ToolTipIcon = “info”
$DB_toolTip.ToolTipTitle = “CreateDB Script”
$DB_toolTip.SetToolTip($bt_DB, “Creating a new database with the neccessary TaxOne data")
$DBFont = New-Object System.Drawing.Font ("Arial", 9, [System.Drawing.FontStyle]::Bold)
$bt_DB.Font = $DBFont

## Create "DB Already exists" Button
$bt_DBex = New-Object "System.Windows.Forms.Button"
$bt_DBex.Text = "Database already exists"
$bt_DBex.SetBounds(340,170,100,22)
$bt_DBex.AutoSize = $true
$DBex_toolTip = New-Object System.Windows.Forms.ToolTip
$DBex_toolTip.AutomaticDelay = 0
$DBex_toolTip.ToolTipIcon = “info”
$DBex_toolTip.ToolTipTitle = “Database already exists”
$DBex_toolTip.SetToolTip($bt_DBex, “If the database is already created in the selected Instance”)

## Create "Set Service Account" Button
$bt_setSVC = New-Object "System.Windows.Forms.Button"
$bt_setSVC.SetBounds(20,210,100,50)
$bt_setSVC.Text = "Set SQL Service Account"
$setSVC_toolTip = New-Object System.Windows.Forms.ToolTip
$setSVC_toolTip.AutomaticDelay = 0
$setSVC_toolTip.ToolTipIcon = “info”
$setSVC_toolTip.ToolTipTitle = “SQL Querys”
$setSVC_toolTip.SetToolTip($bt_setSVC, “Set the Service Account giving him all neccessary permissions.")

## Create "Install IIS 7.5" Button
$bt_iis = New-Object "System.Windows.Forms.Button"
$bt_iis.Text = "Install IIS 7.5"
$bt_iis.SetBounds(140,200,100,25)
$iis_toolTip = New-Object System.Windows.Forms.ToolTip
$iis_toolTip.AutomaticDelay = 0
$iis_toolTip.ToolTipIcon = “info”
$iis_toolTip.ToolTipTitle = “IIS Installation”
$iis_toolTip.SetToolTip($bt_iis, “Install IIS7.5 via Script”)

## Create "TaxOne Modules" Button
$bt_webmod = New-Object "System.Windows.Forms.Button"
$bt_webmod.Text = "Install TaxOne Webmodules"
$bt_webmod.SetBounds(140,210,100,50)
$bt_webmod.AutoSize = $true
$webmod_toolTip = New-Object System.Windows.Forms.ToolTip
$webmod_toolTip.AutomaticDelay = 0
$webmod_toolTip.ToolTipIcon = “info”
$webmod_toolTip.ToolTipTitle = “TaxOne Webmodules”
$webmod_toolTip.SetToolTip($bt_webmod, “Select the Web Modules, that should be installed”)

## Create "Update" Button
$bt_update = New-Object "System.Windows.Forms.Button"
$bt_update.Text = "Update TaxOne Installation"
$bt_update.SetBounds(310,210,100,50)
$bt_update.AutoSize = $true
$update_toolTip = New-Object System.Windows.Forms.ToolTip
$update_toolTip.AutomaticDelay = 0
$update_toolTip.ToolTipIcon = “info”
$update_toolTip.ToolTipTitle = “Update TaxOne Installation”
$update_toolTip.SetToolTip($bt_update, “Starting a script for updating the version of TaxOne.”)

## Create "Exit" Button
$bt_exit = New-Object "System.Windows.Forms.Button"
$bt_exit.Text = "Exit"
$bt_exit.SetBounds(20,270,100,23)
$exit_toolTip = New-Object System.Windows.Forms.ToolTip
$exit_toolTip.AutomaticDelay = 0
$exit_toolTip.ToolTipIcon = “error”
$exit_toolTip.ToolTipTitle = “Exit the Setup”
$exit_toolTip.SetToolTip($bt_exit, “Nope. No way I'm using that.")

# Add "Manual Button" Click Event
## TBD

# Add "Not Installed yet" Click Event
$bt_niy.add_Click({ 
	
	## Starting Setup.exe on debzifsr99
	Write-Log "Not installed yet has been clicked. Opening SQL Setup"
    Start-Process -FilePath "\\debzifsr99\serverinstall$\SQL2008R2\setup.exe"
	sleep 5
	exit
 })
 
# Add "Create" Click Event
$bt_DB.add_Click({
	
	## Get Instance Name for Script	
	$instName = $cb_Instance.get_Text()
	
	## Get DB Name and Create File
	$newDB = $tb_database.get_Text()
		if ($newDB -like "") {
			Write-Host "Please define a database name"
		}
		else {
			$newDBfile = "$newDB" + '.sql'
            $newDBlog = "$newDB" + '_log'
            

			## Creating Script for creating database
			$y = Test-Path $destPathDB 
			if ($y -like $false) {
				mkdir $destPathDB
                Write-Log "Creating $destPathDB" 
			}
			
			Out-File -FilePath "$destPathDB\$newDBfile" -Force
			Set-ItemProperty -Path "$destPathDB\$newDBfile" -Name IsReadOnly -Value $false
			Write-Log "$newDBFile has been created at $destPathDB" 
            sleep -Seconds 3
			Add-Content -Encoding UTF8 -Path "$destPathDB\$newDBfile" -Value "CREATE DATABASE [$newDB]
 CONTAINMENT = NONE
 ON  PRIMARY 
( NAME = N'$newDB', FILENAME = N'D:\Microsoft SQL Server\MSSQL11.$instName\MSSQL\DATA\$newDB.mdf' , SIZE = 5120KB , FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'$newDBlog', FILENAME = N'E:\Microsoft SQL Server\MSSQL11.$instName\MSSQL\DATA\$newDBlog.ldf' , SIZE = 1024KB , FILEGROWTH = 10%)
GO
ALTER DATABASE [$newDB] SET COMPATIBILITY_LEVEL = 110
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
ALTER DATABASE [$newDB] SET READ_COMMITTED_SNAPSHOT OFF 
GO
ALTER DATABASE [$newDB] SET  READ_WRITE 
GO
ALTER DATABASE [$newDB] SET RECOVERY FULL 
GO
ALTER DATABASE [$newDB] SET  MULTI_USER 
GO
ALTER DATABASE [$newDB] SET PAGE_VERIFY CHECKSUM  
GO
ALTER DATABASE [$newDB] SET TARGET_RECOVERY_TIME = 0 SECONDS 
GO
USE [$newDB]
GO
IF NOT EXISTS (SELECT name FROM sys.filegroups WHERE is_default=1 AND name = N'PRIMARY') ALTER DATABASE [$newDB] MODIFY FILEGROUP [PRIMARY] DEFAULT
GO
"	
	} #end of Else

	sleep -Seconds 3
	Write-Log -ForegroundColor DarkGreen "SQL Script successfully created in $destPathDB"
	Write-Log "Name: $newDBfile"
	
	## Creating Database
    Write-Log "Executing SQLCMD Command to Create Database"
    sqlcmd -S $servername\$instName -E -i $destPathDB\$newDBfile 
   

	## Copy TaxOne SQL Querys to local machine
	Write-Log "Copying and executing SQL Querys for TaxOne Installation"
    Write-Host "Copying..."
	Copy-Item -Path "$sourcePathDB" -Destination "$destPathDB" -Recurse -Force -ErrorAction Stop
	sleep -Seconds 3
	Write-Log "Files has been copied from $sourcePathDB to $destPathDB"
    Write-Host "Starting Execution..."
    
    ## Executing TaxOne Querys
    Write-Host "Processing .. 1_Taxportal_SQLSchema.sql"
    sqlcmd -S $servername\$instname -d $newDB -E -i $destPathDB\DB\1_Taxportal_SQLSchema.sql
    Write-Log "$destPathDB\DB\1_Taxportal_SQLSchema.sql executed."

    Write-Host "Processing .. 2_Taxportal_SQLData_1_0_0.sql"
    sqlcmd -S $servername\$instname -d $newDB -E -i $destPathDB\DB\2_Taxportal_SQLData_1_0_0.sql
    Write-Log "$destPathDB\DB\2_Taxportal_SQLData_1_0_0.sql executed."

    Write-Host "Processing .. 3_Taxportal_SQLIndex.sql"
    sqlcmd -S $servername\$instname -d $newDB -E -i $destPathDB\DB\3_Taxportal_SQLIndex.sql
    Write-Log "$destPathDB\DB\3_Taxportal_SQLIndex.sql executed."

    Write-Host "Processing .. TaxPortal_SQLData_TaxCompliance.sql"
    sqlcmd -S $servername\$instname -d $newDB -E -i $destPathDB\DB\TaxPortal_SQLData_TaxCompliance.sql
    Write-Log "$destPathDB\DB\TaxPortal_SQLData_TaxCompliance.sql executed."
    Write-Log "SQL Installation done" 
    Write-Host "SQL Installation done. Please continue by clicking `"Set SQL Service Account`""


})	

# Add "DB already exists" Click Event
$bt_DBex.add_Click({


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
        out-file -FilePath "$destPathDB\$accName.sql"
        Write-Log "$newDBFile has been created at $destPathDB" 
	    Add-Content -Encoding utf8 -Path "$destPathDB\$accName.sql" -Value "USE [master]
GO
CREATE LOGIN [de\$accName] FROM WINDOWS WITH DEFAULT_DATABASE=[master]
GO
USE [$newDB]
GO
CREATE USER [de\$accName] FOR LOGIN [de\$accName]
GO
USE [$newDB]
GO
ALTER ROLE [taxportal_role] ADD MEMBER [de\$accName]
GO"
        sleep -Seconds 3
        sqlcmd -S $servername\$instname -d $newDB -E -i "$destPathDB\$accName.sql"
        Write-Log "$accName successfully allowed on $newDB"
        Write-Host "SQL permissions granted."
        Write-Host "The TaxOne Database Installation is finished. Please continue by install the TaxOne Webmodules."

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

# Add "Install IIS 7.5" Click Event
 $bt_iis.add_Click({
	$w = Test-Path -Path "D:\WebServer\wwwroot\*"
	if ($w -like $true) {
		Write-Host "IIS is already installed on this server"
	}
	else {
	Write-Host "Starting Installation of IIS 7.5..."
	Start-Process -FilePath "\\debzifsr99\serverinstall$\IIS7.x\installIIS7_remote.cmd"
	}
 })

# Add "Tax One Webmodules" Click Event
$bt_webmod.add_Click({
   
    $accName = $global:accName
    $pw = $global:pw
    $instance = $cb_instance.get_Text()
    $destPathWeb = "D:\WebServer\$instance\GeneralWebsite"

    if (Test-Path $destPathWeb) {}
    else {
        mkdir $destPathWeb -Force
    }

	## Create TaxOne Webmodules Election Window
	$webElectWindow = New-Object "System.Windows.Forms.Form"
	$webElectWindow.Size = New-Object System.Drawing.Size @(350,300)
	$webElectWindow.TopMost = $true
	$webElectWindow.Text = "TaxOne Webmodules"
	
	## Create Check Boxes
	$webBox_TAXPORTAL = New-Object System.Windows.Forms.CheckBox
	$webBox_TAXPORTAL.Text = "TaxPortal"
	$webBox_TAXPORTAL.Checked = $true
	$webBox_TAXPORTAL.SetBounds(120,20,100,25)
	$webBox_TAXPORTAL.Enabled = $false

	$webBox_TAXADMIN = New-Object System.Windows.Forms.CheckBox
	$webBox_TAXADMIN.Text = "TaxAdmin"
	$webBox_TAXADMIN.Checked = $true
	$webBox_TAXADMIN.SetBounds(120,50,100,25)
	$webBox_TAXADMIN.Enabled = $false

	$webBox_TIGER = New-Object System.Windows.Forms.CheckBox
	$webBox_TIGER.Text = "Tiger"
	$webBox_TIGER.SetBounds(20,20,100,25)

	$webBox_ZEBRA = New-Object System.Windows.Forms.CheckBox
	$webBox_ZEBRA.Text = "Zebra"
	$webBox_ZEBRA.SetBounds(20,50,100,25)

	$webBox_OCTOPUS = New-Object System.Windows.Forms.CheckBox
	$webBox_OCTOPUS.Text = "Octopus"	
	$webBox_OCTOPUS.SetBounds(20,80,100,25)

	$webBox_PUMA = New-Object System.Windows.Forms.CheckBox
	$webBox_PUMA.Text = "Puma"
	$webBox_PUMA.SetBounds(20,110,100,25)
	
	## Create "Save" Button
	$webBox_save = New-Object "System.Windows.Forms.Button"
	$webBox_save.SetBounds(20,155,100,25)
	$webBox_save.Text = "Start Copy"
	$save_toolTip = New-Object System.Windows.Forms.ToolTip
	$save_toolTip.AutomaticDelay = 0
	$save_toolTip.ToolTipIcon = “info”
	$save_toolTip.ToolTipTitle = “Save”
	$save_toolTip.SetToolTip($webBox_save, “The setup will copy the selected Webmodules to the specfific WebServer directory.")

    ## Create "Update Config" Button
    $webBox_cfg = New-Object System.Windows.Forms.Button
    $webBox_cfg.SetBounds(140,155,130,25)
    $webBox_cfg.Text = "Update Configuration"
    $webBox_cfg.Enabled = $false
    $cfg_toolTip = New-Object System.Windows.Forms.ToolTip
    $cfg_toolTip.AutomaticDelay = 0
	$cfg_toolTip.ToolTipIcon = “info”
	$cfg_toolTip.ToolTipTitle = “Change configuration file”
	$cfg_toolTip.SetToolTip($webBox_cfg, “Updates the configuration file for the selected Webmodules.")

    ## Create "Configure IIS" Button
    $webBox_iis = New-Object System.Windows.Forms.Button
    $webBox_iis.SetBounds(20,190,100,25)
    $webBox_iis.Text = "Configure IIS"
    $webBox_iis.Enabled = $false
    $iis_toolTip = New-Object System.Windows.Forms.ToolTip
    $iis_toolTip.AutomaticDelay = 0
	$iis_toolTip.ToolTipIcon = “info”
	$iis_toolTip.ToolTipTitle = “Configure IIS”
	$iis_toolTip.SetToolTip($webBox_iis, “Converting the folder into applications and set configuration for TaxPortal")

    ## Create "Close" Button
    $webBox_close = New-Object System.Windows.Forms.Button
    $webBox_close.SetBounds(140,190,100,25)
    $webBox_close.Text = "Close"

	
	## Add "Save" Button Function
	$webBox_save.add_Click({

		## TIGER
		if ($webBox_TIGER.Checked -eq $true) {
			Write-Host "TIGER has been selected ..."
			if (Test-Path "$destPathWeb\Tiger")  {
              	Write-Host "TIGER is already installed"
                Write-Log "TIGER is already installed"  
			}
			else {
                Write-Host "Copying TIGER..."
				Copy-Item -Path "$sourcePathWeb\Tiger" -Destination "$destPathWeb\Tiger" -Recurse 
                Write-Log "TIGER copied to $destPathWeb\Tiger"
                # Creating Log Directory and setting ACL
                # mkdir "$destPathWeb\Tiger\Log" -ErrorAction SilentlyContinue
                
			} 			
        }
        else {
			Write-Host "TIGER has not been selected."
			
		}
			
		## ZEBRA
		if ($webBox_ZEBRA.Checked -eq $true) {
			Write-Host "ZEBRA has been selected ..."
			if (Test-Path "$destPathWeb\Zebra")  {
              	Write-Host "ZEBRA is already installed"
                Write-Log "ZEBRA is already installed"  
			}
			else {
                Write-Host "Copying ZEBRA..."
				Copy-Item -Path "$sourcePathWeb\Zebra" -Destination "$destPathWeb\Zebra" -Recurse 
                Write-Log "ZEBRA copied to $destPathWeb\Zebra"
                # Creating Log Directory and setting ACL
                # mkdir "$destPathWeb\Zebra\Log" -ErrorAction SilentlyContinue
			}
        }
		else {
			Write-Host "ZEBRA has not been selected."
		} 
	
			
		## OCTOPUS
		if ($webBox_OCTOPUS.Checked -eq $true) {
			Write-Host "OCTOPUS has been selected..."
			if (Test-Path "$destPathWeb\Octopus")  {
              	Write-Host "OCTOPUS is already installed"
                Write-Log "OCTOPUS is already installed"  
			}
			else {
                Write-Host "Copying OCTOPUS..."
				Copy-Item -Path "$sourcePathWeb\Octopus" -Destination "$destPathWeb\Octopus" -Recurse 
                Write-Log "OCTOPUS copied to $destPathWeb\Octopus"
                # Creating Log Directory and setting ACL
                # mkdir "$destPathWeb\Octopus\Log" -ErrorAction SilentlyContinue

			} 
        }
		else {
			Write-Host "OCTOPUS has not been selected."
		}
			
		## PUMA
		if ($webBox_PUMA.Checked -eq $true) {
			Write-Host "PUMA has been selected..."
			if (Test-Path "$destPathWeb\Puma")  {
              	Write-Host "Puma is already installed"
                Write-Log "Puma is already installed"  
			}
			else {
                Write-Host "Copying PUMA..."
				Copy-Item -Path "$sourcePathWeb\Puma" -Destination "$destPathWeb\Puma" -Recurse 
                Write-Log "PUMA copied to $destPathWeb\Puma"
                # Creating Log Directory and setting ACL
                # mkdir "$destPathWeb\Puma\Log" -ErrorAction SilentlyContinue

			} 
        }
		else {
			Write-Host "PUMA has not been selected."
		}
			
		## TaxPortal
		if (Test-Path "$destPathWeb\TaxPortal")  {
            Write-Host "TaxPortal is already installed"
            Write-Log "TaxPortal is already installed"  
		}
		else {
            Write-Host "Copying TaxPortal..."
			Copy-Item -Path "$sourcePathWeb\TaxPortal" -Destination "$destPathWeb\TaxPortal" -Recurse 
            Write-Log "TaxPortal copied to $destPathWeb\TaxPortal"
            # Creating Log Directory and setting ACL
            # mkdir "$destPathWeb\TaxPortal\Log" -ErrorAction SilentlyContinue

		}
			
		## TaxAdmin
		if (Test-Path "$destPathWeb\TaxAdmin") {
			Write-Host "TaxAdmin is already installed"
            Write-Log "TaxAdmin is already installed" 
		}
		else {
			Write-Host "Copying TaxAdmin..."
			Copy-Item -Path "$sourcePathWeb\TaxAdmin" -Destination "$destPathWeb\TaxAdmin" -Recurse 
            Write-Log "TaxAdmin copied to $destPathWeb\TaxAdmin"
            # Creating Log Directory and setting ACL
            # mkdir "$destPathWeb\TaxAdmin\Log" -ErrorAction SilentlyContinue

		}
    Write-Host "Copying Websources done."
    $webBox_cfg.Enabled = $true
    $webBox_iis.Enabled = $true
		
	})

    ## Add "Update Configuration" Button Function
    $webBox_cfg.add_Click({
 
        $instance = $cb_instance.get_Text()
        $newDB = $tb_database.get_Text()
        $destPathWeb = "D:\WebServer\$instance\GeneralWebsite"
        write-host "Are TaxOne database sources on a different server? (y/n)"
        $a = Read-Host
        if ($a-eq 'y' -or $a -eq "y") {
            write-host "Name: "
            $servername = Read-Host
            Write-log "Changed TaxOne DB Server to $servername"
           
        }
            Write-host "Continueing with $servername"
            Write-Log "Did not changed DB servername."
        
		## TaxPortal
		if (test-path "$destPathWeb\TaxPortal\hibernate.cfg.xml") {
			Write-Host "TaxPortal is already installed. Config won't be touched."
		}
        else { 
            Remove-Item "$destPathWeb\TaxPortal\" -Include "*sample*" -Recurse
            Write-Log "Replacing TaxPortal hibernate.cfg.xml"
           	Out-File -FilePath "$destPathWeb\TaxPortal\hibernate.cfg.xml" -Force -Encoding utf8
			Set-ItemProperty -Path "$destPathWeb\TaxPortal\hibernate.cfg.xml" -Name IsReadOnly -Value $false
			Write-Log "$destPathWeb\TaxPortal\hibernate.cfg.xml has been created"
            sleep -Seconds 3

            Add-Content -Encoding UTF8 -Path "$destPathWeb\TaxPortal\hibernate.cfg.xml" -Value "<?xml version='1.0' encoding='utf-8'?>
<hibernate-configuration xmlns=`"urn:nhibernate-configuration-2.2`">
  <session-factory>
    <property name=`"dialect`">NHibernate.Dialect.MsSql2005Dialect</property>
    <property name=`"connection.provider`">NHibernate.Connection.DriverConnectionProvider</property>
    <property name=`"connection.driver_class`">NHibernate.Driver.SqlClientDriver</property>
    <property name=`"connection.connection_string`">
      Server=$servername\$instance;initial catalog=$newDB;Integrated Security=SSPI;Min Pool Size=1
    </property>
    <property name=`"connection.isolation`">ReadCommitted</property>
  </session-factory>
</hibernate-configuration>
"
		} 

        ## TaxAdmin
		if (test-path "$destPathWeb\TaxAdmin\hibernate.cfg.xml") {
			Write-Host "TaxAdmin is already installed. Config won't be touched."
		}
        else { 
            Remove-Item "$destPathWeb\TaxAdmin\" -Include "*sample*" -Recurse
            Write-Log "Replacing TaxAdmin hibernate.cfg.xml"
           	Out-File -FilePath "$destPathWeb\TaxAdmin\hibernate.cfg.xml" -Force -Encoding utf8
			Set-ItemProperty -Path "$destPathWeb\TaxAdmin\hibernate.cfg.xml" -Name IsReadOnly -Value $false
			Write-Log "$destPathWeb\TaxAdmin\hibernate.cfg.xml has been created"
            sleep -Seconds 3

            Add-Content -Encoding utf8 -Path "$destPathWeb\TaxAdmin\hibernate.cfg.xml" -Value "<?xml version='1.0' encoding='utf-8'?>
<hibernate-configuration xmlns=`"urn:nhibernate-configuration-2.2`">
  <session-factory>
    <property name=`"dialect`">NHibernate.Dialect.MsSql2005Dialect</property>
    <property name=`"connection.provider`">NHibernate.Connection.DriverConnectionProvider</property>
    <property name=`"connection.driver_class`">NHibernate.Driver.SqlClientDriver</property>
    <property name=`"connection.connection_string`">
      Server=$servername\$instance;initial catalog=$newDB;Integrated Security=SSPI;Min Pool Size=1
    </property>
    <property name=`"connection.isolation`">ReadCommitted</property>
  </session-factory>
</hibernate-configuration>
"
		} 
        ## TIGER
		if ($webBox_TIGER.Checked -eq $true) {
			Write-Host "TIGER has been selected ..."
			Remove-Item "$destPathWeb\Tiger\" -Include "*sample*" -Recurse
            Write-Log "Replacing TIGER hibernate.cfg.xml"
           	Out-File -FilePath "$destPathWeb\Tiger\hibernate.cfg.xml" -Force -Encoding utf8
			Set-ItemProperty -Path "$destPathWeb\Tiger\hibernate.cfg.xml" -Name IsReadOnly -Value $false
			Write-Log "$destPathWeb\Tiger\hibernate.cfg.xml has been created"
            sleep -Seconds 3

            Add-Content -Encoding utf8 -Path "$destPathWeb\Tiger\hibernate.cfg.xml" -Value "<?xml version='1.0' encoding='utf-8'?>
<hibernate-configuration xmlns=`"urn:nhibernate-configuration-2.2`">
  <session-factory>
    <property name=`"dialect`">NHibernate.Dialect.MsSql2005Dialect</property>
    <property name=`"connection.provider`">NHibernate.Connection.DriverConnectionProvider</property>
    <property name=`"connection.driver_class`">NHibernate.Driver.SqlClientDriver</property>
    <property name=`"connection.connection_string`">
      Server=$servername\$instance;initial catalog=$newDB;Integrated Security=SSPI;Min Pool Size=1
    </property>
    <property name=`"connection.isolation`">ReadCommitted</property>
  </session-factory>
</hibernate-configuration>
"
		} 

        ## PUMA
		if ($webBox_PUMA.Checked -eq $true) {
			Write-Host "Puma has been selected ..."
			Remove-Item "$destPathWeb\Puma\" -Include "*sample*" -Recurse
            Write-Log "Replacing PUMA hibernate.cfg.xml"
           	Out-File -FilePath "$destPathWeb\Puma\hibernate.cfg.xml" -Force -Encoding utf8
			Set-ItemProperty -Path "$destPathWeb\Puma\hibernate.cfg.xml" -Name IsReadOnly -Value $false
			Write-Log "$destPathWeb\Puma\hibernate.cfg.xml has been created"
            sleep -Seconds 3

            Add-Content -Encoding utf8 -Path "$destPathWeb\Puma\hibernate.cfg.xml" -Value "<?xml version='1.0' encoding='utf-8'?>
<hibernate-configuration xmlns=`"urn:nhibernate-configuration-2.2`">
	<session-factory>
        <property name=`"dialect`">NHibernate.Dialect.MsSql2005Dialect</property>
        <property name=`"connection.provider`">NHibernate.Connection.DriverConnectionProvider</property>
        <property name=`"connection.driver_class`">NHibernate.Driver.SqlClientDriver</property>
		<property name=`"connection.connection_string`">
			Server=$servername\$instance;initial catalog=$newDB;Integrated Security=SSPI;Min Pool Size=1
		</property>
		<property name=`"connection.isolation`"ReadCommitted</property>
	</session-factory>
</hibernate-configuration>
"
		} 

        ## ZEBRA	
		if ($webBox_ZEBRA.Checked -eq $true) {
			Write-Host "ZEBRA has been selected ..."
			Remove-Item "$destPathWeb\Zebra\" -Include "*sample*" -Recurse
            Write-Log "Replacing ZEBRA hibernate.cfg.xml"
           	Out-File -FilePath "$destPathWeb\Zebra\hibernate.cfg.xml" -Force -Encoding utf8
			Set-ItemProperty -Path "$destPathWeb\Zebra\hibernate.cfg.xml" -Name IsReadOnly -Value $false
			Write-Log "$destPathWeb\Zebra\hibernate.cfg.xml has been created"
            sleep -Seconds 3

            Add-Content -Encoding utf8 -Path "$destPathWeb\Zebra\hibernate.cfg.xml" -Value "<?xml version='1.0' encoding='utf-8'?>
<hibernate-configuration xmlns=`"urn:nhibernate-configuration-2.2`">
	<session-factory>
        <property name=`"dialect`">NHibernate.Dialect.MsSql2005Dialect</property>
        <property name=`"connection.provider`">NHibernate.Connection.DriverConnectionProvider</property>
        <property name=`"connection.driver_class`">NHibernate.Driver.SqlClientDriver</property>
		<property name=`"connection.connection_string`">
			Server=$servername\$instance;initial catalog=$newDB;Integrated Security=SSPI;Min Pool Size=1
		</property>
		<property name=`"connection.isolation`"ReadCommitted</property>
	</session-factory>
</hibernate-configuration>
"
		} 	

        ## OCTOPUS
		if ($webBox_OCTOPUS.Checked -eq $true) {
			Write-Host "OCTOPUS has been selected ..."
			Remove-Item "$destPathWeb\Octopus\" -Include "*sample*" -Recurse
            Write-Log "Replacing OCTOPUS hibernate.cfg.xml"
           	Out-File -FilePath "$destPathWeb\Octopus\hibernate.cfg.xml" -Force -Encoding utf8
			Set-ItemProperty -Path "$destPathWeb\Octopus\hibernate.cfg.xml" -Name IsReadOnly -Value $false
			Write-Log "$destPathWeb\Octopus\hibernate.cfg.xml has been created"
            sleep -Seconds 3

            Add-Content -Encoding utf8 -Path "$destPathWeb\Octopus\hibernate.cfg.xml" -Value "<?xml version='1.0' encoding='utf-8'?>
<hibernate-configuration xmlns=`"urn:nhibernate-configuration-2.2`">
	<session-factory>
        <property name=`"dialect`">NHibernate.Dialect.MsSql2005Dialect</property>
        <property name=`"connection.provider`">NHibernate.Connection.DriverConnectionProvider</property>
        <property name=`"connection.driver_class`">NHibernate.Driver.SqlClientDriver</property>
		<property name=`"connection.connection_string`">
			Server=$servername\$instance;initial catalog=$newDB;Integrated Security=SSPI;Min Pool Size=1
		</property>
		<property name=`"connection.isolation`"ReadCommitted</property>
	</session-factory>
</hibernate-configuration>
"		
        }
        Write-Host "Configuration files have been updated. Continue by Configuring the IIS"

    })

    ## Add "Configure IIS" Button Function
    $webBox_iis.add_Click({
        $x = $appcmd.ToString()
        ## Create Application Pool
        Write-Host "Starting IIS Configuration. View Log for further information"
        start "$x" -ArgumentList "add apppool -name:`"$instance`""
        Write-Log "Application Pool created: $instance"
        sleep -Seconds 5

        ## Change ApplicationPool in .NET 4.0 and Classic Pipeline Mode
        start "$x" -ArgumentList "set apppool `"$instance`" /managedPipelineMode:Classic"
        sleep -Seconds 5
        start "$x" -ArgumentList "set apppool `"$instance`" /managedRuntimeVersion:v4.0"
        sleep -Seconds 5

        ## set bindings
        start "$x" -ArgumentList "add site -name:`"$instance`" -physicalpath:`"$destPathWeb`"" 
            #start "$x" -ArgumentList "add site -name:`"$instance`" -bindings:`"HTTP/*:80:$instance`" -physicalpath:`"$destPathWeb`"" 
            #Write-Log "Creating website '$instance' with binding 'HTTP/*:80:$instance' and at $destPathWeb"
        sleep -Seconds 2
        start "$x" -ArgumentList "set site `"$instance`" /+bindings.[protocol='http',bindingInformation='*:80:']"
            #start "$x" -ArgumentList "set site `"$instance`" /+bindings.[protocol='http',bindingInformation='*:80:$instance.de.kworld.kpmg.com']"
            #Write-Log "Set Bindings: [protocol='http',bindingInformation='*:80:$instance.de.kworld.kpmg.com']"
        sleep -Seconds 2

        ## Set web site to use new application pool
        start "$x" -ArgumentList "set config -section:system.applicationHost/sites -[name=`'$instance`'].applicationDefaults.applicationPool:`"$instance`"  -commit:apphost"
        sleep -Seconds 2
        start "$x" -ArgumentList "set config -section:system.applicationHost/applicationPools -[name=`'$instance`'].processModel.identityType:`"SpecificUser`" -commit:apphost"
        sleep -Seconds 2

        ## Set application pool identity
        start "$x" -ArgumentList "set config -section:system.applicationHost/applicationPools -[name=`'$instance`'].processModel.userName:`"de\$accName`" -commit:apphost"
        sleep -Seconds 2
        start "$x" -ArgumentList "set config -section:system.applicationHost/applicationPools -[name=`'$instance`'].processModel.password:`"$pw`" -commit:apphost"

        ## Set IIS log directory
        mkdir "E:\LogFiles\$instance" -ErrorAction SilentlyContinue
        start "$x" -ArgumentList "set config -section:system.applicationHost/sites -[name=`'$instance`'].logfile.directory:`"E:\LogFiles\$instance`""

        
        ## Convert virtual directories to applications
        Write-Host "Converting virtual directories to applications..."
        if (Test-Path "$destPathWeb\TaxPortal") {
            start "$x" -ArgumentList "add app /site.name:`"$instance`" /path:/`"TaxPortal`" /physicalPath:`"$destPathweb\TaxPortal`""
            Write-Log "Converting TaxPortal to application"
            sleep -Seconds 2
        }
        if (Test-Path "$destPathWeb\TaxAdmin") {
            start "$x" -ArgumentList "add app /site.name:`"$instance`" /path:/`"TaxAdmin`" /physicalPath:`"$destPathweb\TaxAdmin`""
            Write-Log "Converting TaxAdmin to application"
            sleep -Seconds 2
        }
        if (Test-Path "$destPathWeb\Tiger") {
            start "$x" -ArgumentList "add app /site.name:`"$instance`" /path:/`"Tiger`" /physicalPath:`"$destPathweb\Tiger`""
            Write-Log "Converting Tiger to application"
            sleep -Seconds 2
        }
        if (Test-Path "$destPathWeb\Zebra") {
            start "$x" -ArgumentList "add app /site.name:`"$instance`" /path:/`"Zebra`" /physicalPath:`"$destPathweb\Zebra`""
            Write-Log "Converting Zebra to application"
            sleep -Seconds 2
        }
        if (Test-Path "$destPathWeb\Octopus") {
            start "$x" -ArgumentList "add app /site.name:`"$instance`" /path:/`"Octopus`" /physicalPath:`"$destPathweb\Octopus`""
            Write-Log "Converting Octopus to application"
            sleep -Seconds 2
        }
        if (Test-Path "$destPathWeb\Puma") {
            start "$x" -ArgumentList "add app /site.name:`"$instance`" /path:/`"Puma`" /physicalPath:`"$destPathweb\Puma`""
            Write-Log "Converting Puma to application"
            sleep -Seconds 2
        }
        if (Test-Path "$destPathWeb\ETax") {
            start "$x" -ArgumentList "add app /site.name:`"$instance`" /path:/`"ETax`" /physicalPath:`"$destPathweb\Etax`""
            Write-Log "Converting ETax to application"
            sleep -Seconds 2
        }
        Write-Host "Done"
        ## Set TaxPortal Default Document
        start "$x" -ArgumentList "set config `"$instance/TaxPortal`" /section:defaultDocument `"/+files.[@start,value='TaxPortal.aspx']`" /commit:`"$instance/TaxPortal`""
        Write-Log "Setting TaxPortal Default Document"


             
    })

    ## Add "Close" Button Function
    $webBox_close.add_Click({
        $webElectWindow.Close()
    })
	
	$webElectWindow.Controls.Add($webBox_TIGER)
	$webElectWindow.Controls.Add($webBox_ZEBRA)
	$webElectWindow.Controls.Add($webBox_OCTOPUS)
	$webElectWindow.Controls.Add($webBox_PUMA)
	$webElectWindow.Controls.Add($webBox_save)
	$webElectWindow.Controls.Add($webBox_TAXPORTAL)
	$webElectWindow.Controls.Add($webBox_TAXADMIN)
    $webElectWindow.Controls.Add($webBox_cfg)
    $webElectWindow.Controls.Add($webBox_iis)
    $webElectWindow.Controls.Add($webBox_close)
	$webElectWindow.ShowDialog()
	
	
})

# Add "Update" Click Event
$bt_update.add_Click({
    & '\\defr2app31\d$\TaxOne AutoInstall\TaxOneUpdate.ps1'
})

# Add "Exit" Click Event
$bt_exit.add_Click({
	[System.Windows.Forms.MessageBox]::Show("That button has no function at all. It is just there. Deal with it.", "Nope.", 0, [System.Windows.Forms.MessageBoxIcon]::Error)
})

## Add Objects to Main Window
$mainForm.Controls.Add($cb_Instance)
#$mainForm.Controls.Add($bt_niy)
$mainForm.Controls.Add($tb_database)
$mainForm.Controls.Add($bt_DBex)
$mainForm.Controls.Add($hint)
$mainForm.Controls.Add($hint2)
$mainForm.Controls.Add($hint3)
$mainForm.Controls.Add($header)
$mainForm.Controls.Add($bt_DB)
$mainForm.Controls.Add($bt_setSVC)
#$mainForm.Controls.Add($bt_iis)
$mainForm.Controls.Add($bt_update)
$mainForm.Controls.Add($bt_webmod)
$mainForm.Controls.Add($bt_exit)
#$mainForm.Controls.Add($bt_Manual)
$mainForm.ShowDialog()


