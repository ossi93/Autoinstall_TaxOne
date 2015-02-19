
### UPDATESCRIPT FOR TAXONE SQL-DATABASES

## 19.01.2015 - Thomas Scholz & Robert Ostwald 

## Version: 1.0


#########################################################################

#### GUI-CODE:

### GENERAL:

## LOADING ASSEMBLY FOR WINDOWS FORMS
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

## FONTS
$headlineFont = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Underline)
$font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::regular)
$underlined = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Underline)

## CREATING A NEW WINDOW 
$window = New-Object System.Windows.Forms.Form
$windowY = 360
$window.Size = New-Object System.Drawing.Size @(470,$windowY)
$window.Text = “TaxOne - Update-Script”
$window.StartPosition = “CenterScreen”
## STATIC WINDOWSIZE:
$window.FormBorderStyle = 'FixedDialog'
## $window.BackgroundImage = $Image

## CREATE A HEADLINE
$headline = New-Object System.Windows.Forms.Label
$headline.Text = “TaxOne - Update-Script”
$headline.AutoSize = $True
$headline.Location = New-Object System.Drawing.Size(20,20)
$headline.Font = $HeadlineFont

#########################################################################

### UPDATEPATH:

## CREATE A LABEL FOR THE UPDATEPATH
$UpdatepathDBLabel = New-Object System.Windows.Forms.Label
$UpdatepathDBLabel.Text = "Update-path:"
$UpdatepathDBLabel.AutoSize = $True
$UpdatepathDBLabel.Location = New-Object System.Drawing.Size(20,80)
$UpdatepathDBLabel.Font = $underlined

## CREATE A TEXTBOX FOR THE UPDATEPATH
$UpdatepathDBBox = New-Object System.Windows.Forms.TextBox
$UpdatepathDBBox.SetBounds(170,80,180,100)

## CREATE A TOOLTIP FOR THE UPDATEPATH-TEXTBOX
$TTUpdatepathDBLabel = New-Object System.Windows.Forms.ToolTip
$TTUpdatepathDBLabel.SetToolTip($UpdatepathDBLabel, "Add the Path, where all SQL-Scripts are located.")
$TTUpdatepathDBLabel.ToolTipIcon = 'info'

# CREATE A BROWSEBUTTON FOR THE UPDATEPATH
$BrowseUpdatepathButton = New-Object System.Windows.Forms.Button
$BrowseUpdatepathButton.Location = New-Object System.Drawing.Size(355,78)
$BrowseUpdatepathButton.Size = New-Object System.Drawing.Size(75,24)
$BrowseUpdatepathButton.Text = “Browse”
$BrowseUpdatepathButton.Font = $Font


#########################################################################

### INSTANCES:

## CREATE A LABEL FOR THE MULTIPLE INSTANCE QUERY
$multiInst = New-Object System.Windows.Forms.Label
$multiInst.Text = "Instance:"
$multiInst.AutoSize = $True
$multiInst.Location = New-Object System.Drawing.Size(20,120)
$multiInst.Font = $underlined 

## CREATE A CHECKBUTTON FOR THE MULTIPLE INSTANCE QUERY
$multiInstCheck = New-Object System.Windows.Forms.CheckBox
$multiInstCheck.Text = "Multiple instances"
$multiInstCheck.SetBounds(170,112,300,40)
$multiInstCheck.Checked = $false
$multiInstCheck.Font = $font

## CREATE A TOOLTIP FOR THE MULTIPLE INSTANCE QUERY
$TTmultiInst = New-Object System.Windows.Forms.ToolTip
$TTmultiInst.SetToolTip($multiInst, "Check, if there is more than one instance.")
$TTmultiInst.ToolTipIcon = 'info'

## CREATE A LABEL FOR THE INSTANCE-NAME
$instNameLabel = New-Object System.Windows.Forms.Label
$instNameLabel.Text = "Name:"
$instNameLabel.AutoSize = $True
$instNameY = 160
$instNameLabel.Location = New-Object System.Drawing.Size(60,$instNameY)
$instNameLabel.Font = $font

## CREATE A COMBOBOX FOR THE INSTANCE-NAME
$instNameBox = New-Object System.Windows.Forms.ComboBox

## CREATE AN ARRAY WITH ALL INSTANCES INSTALLED ON THE SERVER
$instNameArray = (Get-Itemproperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances
$i = 0

## SHOW ALL INSTANCES IN THE COMBOBOX
for ($i; $i -lt @($instNameArray).length; $i++) {
	[void] $instNameBox.Items.Add($instNameArray[$i])	
	}
$instNameBox.SetBounds(170,160,180,100)

## CREATE A TOOLTIP FOR THE INSTANCENAME-TEXTBOX
$TTinstNameLabel= New-Object System.Windows.Forms.ToolTip
$TTinstNameLabel.SetToolTip($instNameLabel, "Select the name of the instance of all databases.")
$TTinstNameLabel.ToolTipIcon = 'info'

# CREATE A SET FOR THE DBARRAY
$SetButton = New-Object System.Windows.Forms.Button
$SetButton.Location = New-Object System.Drawing.Size(355,159)
$SetButton.Size = New-Object System.Drawing.Size(75,24)
$SetButton.Text = “Set”
$SetButton.Font = $Font


##########################################################################

###THESE OBJECTS ARE ONLY VISIBLE WHEN THE INSTANCE CHECKBOX IS ACTIVATED

## CREATE A LABEL FOR THE QUANTITY OF INSTANCES
$instQuantityLabel = New-Object System.Windows.Forms.Label
$instQuantityLabel.Text = "Quantity"
$instQuantityLabel.AutoSize = $True
$instQuantityLabel.Location = New-Object System.Drawing.Size(60,160)
$instQuantityLabel.Font = $font

## CREATE A TEXTBOX FOR THE QUANTITY OF INSTANCES
$instQuantityBox = New-Object System.Windows.Forms.TextBox
$instQuantityBox.SetBounds(170,160,260,100)

## CREATE A TOOLTIP FOR THE INSTANCEQUANTITY-TEXTBOX
$TTinstQuantityLabel = New-Object System.Windows.Forms.ToolTip
$TTinstQuantityLabel.SetToolTip($instQuantityLabel, "Enter the quantity of instances you want to update.")
$TTinstQuantityLabel.ToolTipIcon = 'info'

## CREATE A LABEL FOR THE FIRST INSTANCE-NUMBER
$instNumLabel = New-Object System.Windows.Forms.Label
$instNumLabel.Text = "First Number:"
$instNumLabel.AutoSize = $True
$instNumLabel.Location = New-Object System.Drawing.Size(60,200)
$instNumLabel.Font = $font

## CREATE A TEXTBOX FOR THE INSTANCE-NUMBER
$instNumBox = New-Object System.Windows.Forms.TextBox
$instNumBox.SetBounds(170,200,260,100)

## CREATE A TOOLTIP FOR THE INSTANCENUMBER-TEXTBOX
$TTinstNumLabel = New-Object System.Windows.Forms.ToolTip
$TTinstNumLabel.SetToolTip($instNumLabel, "Enter the number of the first instance you want to update. Ignore zeros`ne.q. TaxOneGroup01:`n1 NOT 01")
$TTinstNumLabel.ToolTipIcon = 'info'

#########################################################################

### DATABASES:

## CREATE A DATABASE-HEADLINE
$DBHeadLine = New-Object System.Windows.Forms.Label
$DBHeadLine.Text = "Database:"
$DBHeadLine.AutoSize = $True
$DBheadlineY = 200
$DBHeadLine.Location = New-Object System.Drawing.Size(20,$DBHeadLineY)
$DBHeadLine.Font = $underlined

## CREATE A CHECKBUTTON FOR THE MULTIPLE DATABASE QUERY
$multiDBCheck = New-Object System.Windows.Forms.CheckBox
$multiDBCheck.Text = "Multiple databases"
$multiDBCheckY = 192
$multiDBCheck.SetBounds(170,$multiDBCheckY,300,40)
$multiDBCheck.Checked = $false
$multiDBCheck.Font = $font

## CREATE A TOOLTIP FOR THE MULTIPLE INSTANCE QUERY
$TTDBHeadLine = New-Object System.Windows.Forms.ToolTip
$TTDBHeadLine.SetToolTip($DBHeadLine, "Check, if there is more than one database.")
$TTDBHeadLine.ToolTipIcon = 'info'

## CREATE A LABEL FOR THE DATABASE-NAME
$DBNameLabel = New-Object System.Windows.Forms.Label
$DBNameLabel.Text = "Name:"
$DBNameLabel.AutoSize = $True
$DBNameY = 240
$DBNameLabel.Location = New-Object System.Drawing.Size(60,$DBNameY)
$DBNameLabel.Font = $font

### CREATE A TEXTBOX FOR THE DATABASE-NAME
#$DBNameBox = New-Object System.Windows.Forms.TextBox
#$DBNameBox.SetBounds(170,$DBNameY,260,100)
## CREATE A TOOLTIP FOR THE DBNAME-TEXTBOX
#
$TTDBNameBox = New-Object System.Windows.Forms.ToolTip
$TTDBNameBox.SetToolTip($DBNameBox, "Enter the name of the instance of all databases`ne.g. Taxportal01")
$TTDBNameBox.ToolTipIcon = 'info'

## CREATE A COMBOBOX FOR THE INSTANCE-NAME
$DBNameBox = New-Object System.Windows.Forms.ComboBox
$DBNameBox.SetBounds(170,$DBNameY,260,100)

## CREATE A TOOLTIP FOR THE INSTANCENAME-TEXTBOX
$TTDBNameLabel= New-Object System.Windows.Forms.ToolTip
$TTDBNameLabel.SetToolTip($DBNameLabel, "Select the name of the database.")
$TTDBNameLabel.ToolTipIcon = 'info'

##########################################################################################

		
###THESE OBJECTS ARE ONLY VISIBLE IF THE DB CHECKBOX IS ACTIVATED

## CREATING A TEXTBOX INSTNAMEBOX
$instNameBox2 = New-Object System.Windows.Forms.TextBox
$instNameBox2.SetBounds(170,(160+$y),260,100)

## CREATING A TEXTBOX DBNAMEBOX
$DBNameBox2 = New-Object System.Windows.Forms.TextBox
$DBNameBox2.SetBounds(170,($DBNameY+$y),260,100)

## CREATE A LABEL FOR THE QUANTITY OF DATABASES
$DBQuantityLabel = New-Object System.Windows.Forms.Label
$DBQuantityLabel.Text = "Quantity:"
$DBQuantityLabel.AutoSize = $True
$DBQuantityY = 160
$DBQuantityLabel.Location = New-Object System.Drawing.Size(60,($DBQuantityY+$y))
$DBQuantityLabel.Font = $font

## CREATE A TEXTBOX FOR THE QUANTITY OF DATABASES
$DBQuantityBox = New-Object System.Windows.Forms.TextBox
$DBQuantityBox.SetBounds(170,($DBQuantityY),260,100)

## CREATE A TOOLTIP FOR THE DBQUANTITY-TEXTBOX
$TTDBQuantityBox = New-Object System.Windows.Forms.ToolTip
$TTDBQuantityBox.SetToolTip($DBQuantityLabel, "Enter the quantity of databases you want to update.")
$TTDBQuantityBox.ToolTipIcon = 'info'

## CREATE A LABEL FOR THE DATABASE-NUMBER
$DBNumLabel = New-Object System.Windows.Forms.Label
$DBNumLabel.Text = "First Number:"
$DBNumLabel.AutoSize = $True
$DBNumLabelY = 200
$DBNumLabel.Location = New-Object System.Drawing.Size(60,($DBNumLabelY+$y))
$DBNumLabel.Font = $font

## CREATE A TEXTBOX FOR THE DATABASE-NUMBER
$DBNumBox = New-Object System.Windows.Forms.TextBox
$DBNumBox.SetBounds(170,($DBNumLabelY+$y),260,100)

## CREATE A TOOLTIP FOR THE DBNUMBER-TEXTBOX
$TTDBNumBox = New-Object System.Windows.Forms.ToolTip
$TTDBNumBox.SetToolTip($DBNumLabel, "Enter the number of the first database you want to update.Ignore zeros`ne.q. Taxportal01:`n1 NOT 01")
$TTDBNumBox.ToolTipIcon = 'info'

##########################################################################################

# UPDATEBUTTON
$update = New-Object System.Windows.Forms.Button
$button = 280
$update.Location = New-Object System.Drawing.Size(20,$button)
$update.Size = New-Object System.Drawing.Size(75,23)
$update.Text = “Update”
$update.Font = $Font

$TTupdate = New-Object System.Windows.Forms.ToolTip
$TTupdate.SetToolTip($update, "Update will execute your input.")
$TTupdate.ToolTipIcon = 'info'

# EXITBUTTON
$exit = New-Object System.Windows.Forms.Button
$exit.Location = New-Object System.Drawing.Size(355,$button)
$exit.Size = New-Object System.Drawing.Size(75,23)
$exit.Text = “Exit”
$exit.Font = $Font

$browseUpdatePath = New-Object System.Windows.Forms.FolderBrowserDialog
$browseUpdatePath.ShowNewFolderButton = $false
$browseUpdatePath.Description = "Choose a directory"
$browseUpdatePath.$Font

#######################################################################

### CLICK-EVENTS 

## CLICKEVENT FOR THE BROWSE BUTTON

$browseUpdatePathbutton.Add_Click({

	$browseUpdatePath.ShowDialog()
	
	## RESETTING THE UPDATEPATH
	$updatepathDBBox.set_Text("")
	
	$a = $browseUpdatePath.SelectedPath
	
	$updatepathDBBox.AppendText($a)
	
})

#######################################################################

## CLICKEVENT FOR THE SET-BUTTON

# SERVERNAME
$serverName = gc env:computerName 

#CLICKEVENT FOR THE SET-BUTTON 
$SetButton.add_Click({
	
	##CLEAR OLD INFORMATION OF THE DBNAMEBOX
	$DBNameBox.Items.Clear()
	$DBNameArray = $null
	
	##DEFINING A INSTANCE-PATH FOR THE QUERY
	$instNameBoxText = $instNameBox.Text
	$dbpath = "$serverName\$instNameBoxText"

	## CREATE AN ARRAY WITH ALL INSTANCES INSTALLED ON THE SERVER
	$DBNameArray = Invoke-Sqlcmd -Query " select DATABASE_NAME   = db_name(s_mf.database_id) from sys.master_files s_mf where s_mf.state = 0 and has_dbaccess(db_name(s_mf.database_id)) = 1 group by s_mf.database_id order by 1" -ServerInstance $dbpath | select -Expand DATABASE_NAME

	$i = 0
	## SHOW ALL INSTANCES IN THE COMBOBOX
	for ($i; $i -lt @($DBNameArray).length; $i++) {
		[void] $DBNameBox.Items.Add($DBNameArray[$i])	
		}
		
	$window.Controls.Add($DBNameBox)
	$window.Refresh()
})

## CLICKEVENT FOR THE INSTANCE CHECKBOX 
$multiInstCheck.add_CheckedChanged({

	## IF THE INSTACE CHECKBOX IS ACTIVATED
	If ($multiInstCheck.Checked -eq $true){
	
		## ALTERING THE INSTANCE-LABEL
		$multiInst.Text = "Instances:"
	
		$y = 80
		$y2 = 80
	
		If ($multiDBCheck.Checked -eq $true){
		
			$y = 80
			$y2 = 160
		
		}
		
		## SETTING A NEW HIGH OF ALL OBJECTS THAT ARE UNDER THE CHECKBOX LOCATED 
		$window.Size = New-Object System.Drawing.Size @(470,($windowY+$y2))
		$instNameLabel.Location = New-Object System.Drawing.Size(60,(160+$y))
		
		### ALTERING THE TYPE-NAME OF INSTNAMEBOX
		## DELETING THE COMBOBOX
		$instNameBox.Visible = $false
		$SetButton.Visible = $false
		$instNameBox.Text = ""
		$instNameBox2.Visible = $true
		$instNameBox2.SetBounds(170,(160+$y),260,100)
		$DBheadline.Location = New-Object System.Drawing.Size(20,($DBheadlineY+$y))
		$multiDBCheck.SetBounds(170,($multiDBCheckY+$y),300,40)
		$DBQuantityLabel.Location = New-Object System.Drawing.Size(60,($DBQuantityY+$y2))
		$DBQuantityBox.SetBounds(170,($DBQuantityY+$y2),260,100)
		$DBNumLabel.Location = New-Object System.Drawing.Size(60,($DBNumLabelY+$y2))
		$DBNumBox.SetBounds(170,($DBNumLabelY+$y2),260,100)
		$DBNameLabel.Location = New-Object System.Drawing.Size(60,($DBNameY+$y2))
		$DBNameBox.SetBounds(170,($DBNameY+$y2),260,100)
		$DBNameBox2.SetBounds(170,($DBNameY+$y2),260,100)
		$update.Location = New-Object System.Drawing.Size(20,($button+$y2))
		$exit.Location = New-Object System.Drawing.Size(355,($button+$y2))
		
		## ADD THE NEW OBJECTS TO THE WINDOW
		$window.Controls.Add($Setbutton)
		$window.Controls.Add($instNameBox2)
		$window.Controls.Add($instQuantityLabel)
		$window.Controls.Add($instQuantityBox)
		$window.Controls.Add($instNumLabel)
		$window.Controls.Add($instNumBox)
		
		## ALTER THE TOOLTIP FOR THE INSTANCENAME
	    $TTinstNameLabel.SetToolTip($instNameLabel, "Enter the main name of the Instance`ne.q. TaxOneGroup01; TaxOneGroup02...TaxOneGroupXY:`nTaxOneGroup")
		
		$window.Refresh()
		
	}
		
	## IF THE INSTANCECHECKBOX IS NOT ACTIVATED	
	Else { 
	
	## ALTERING THE INSTANCE-LABEL
	$multiInst.Text = "Instance:"
	
	$y = 0
	
		If ($multiDBCheck.Checked -eq $true){$y = 80}
		
		## SETTING THE HIGH OF ALL OBJECTS THAT ARE UNDER THE CHECKBOX LOCATED 
		$window.Size = New-Object System.Drawing.Size @(470,($windowY+$y))
		$instNameLabel.Location = New-Object System.Drawing.Size(60,(160))
		$instNameBox2.Visible = $false
		$SetButton.Visible = $true
		$instNameBox.Visible = $true
		$instNameBox.SetBounds(170,(160),180,100)
		$DBheadline.Location = New-Object System.Drawing.Size(20,($dbheadlineY))
		$multiDBCheck.SetBounds(170,($multiDBCheckY),300,40)
		$DBQuantityLabel.Location = New-Object System.Drawing.Size(60,($DBQuantityY+$y))
		$DBQuantityBox.SetBounds(170,($DBQuantityY+$y),260,100)
		$DBNumLabel.Location = New-Object System.Drawing.Size(60,($DBNumLabelY+$y))
		$DBNumBox.SetBounds(170,($DBNumLabelY+$y),260,100)
		$DBNameLabel.Location = New-Object System.Drawing.Size(60,($DBNameY+$y))
		$DBNameBox.SetBounds(170,($DBNameY+$y),260,100)
		$update.Location = New-Object System.Drawing.Size(20,($button+$y))
		$exit.Location = New-Object System.Drawing.Size(355,($button+$y))
		
		$TTinstNameLabel.SetToolTip($instNameLabel, "Enter the Name of the instance of all databases.")
		
		## DELETING THE NEW OBJECTS
		$window.Controls.Add($SetButton)
		$window.Controls.Add($instNameBox)
		$window.Controls.Add($instNameBox2)
		$window.Controls.Remove($instQuantityLabel)
		$window.Controls.Remove($instQuantityBox)
		$window.Controls.Remove($instNumLabel)
		$window.Controls.Remove($instNumBox)

		
		$window.Refresh()
	}
})

#######################################################################

## CLICKEVENT FOR THE DBCHECKBOX 
$multiDBCheck.add_CheckedChanged({

	## ALTERING THE DB-LABEL
	$DBHeadLine.Text = "Databases:"
	
	## IF THE DB CHECKBOX IS ACTIVATED
	If ($multiDBCheck.Checked -eq $true){	
	
	$y = 80
	
		If ($multiInstCheck.Checked -eq $true){	
		 
		$y = $y + 80}
	 
		
		## SETTING A NEW HIGH OF ALL OBJECTS THAT ARE UNDER THE CHECKBOX LOCATED 
	    
		$window.Size = New-Object System.Drawing.Size @(470,($windowY+$y))
		$DBNameLabel.Location = New-Object System.Drawing.Size(60,(240+$y))
		
		### ALTERING THE TYPE-NAME OF DBNAMEBOX
		## DELETING THE COMBOBOX
		$DBNameBox.Visible = $false
		$DBNameBox2.Visible = $true
		$DBNameBox.Text = ""
		$DBNameBox2.SetBounds(170,($DBNameY+$y),260,100)
		$update.Location = New-Object System.Drawing.Size(20,($button+$y))
		$exit.Location = New-Object System.Drawing.Size(355,($button+$y))
		$DBQuantityBox.SetBounds(170,($DBQuantityY+$y),260,100)
		$DBQuantityLabel.Location = New-Object System.Drawing.Size(60,($DBQuantityY+$y))
		$DBNumLabel.Location = New-Object System.Drawing.Size(60,($DBNumLabelY+$y))
		$DBNumBox.SetBounds(170,($DBNumLabelY+$y),260,100)

		## ADD THE NEW OBJECTS TO THE WINDOW
		$window.Controls.Add($DBNameBox2)
		$window.Controls.Add($DBQuantityLabel)
		$window.Controls.Add($DBQuantityBox)
		$window.Controls.Add($DBNumLabel)
		$window.Controls.Add($DBNumBox)
		
		## ALTER A TOOLTIP FOR THE DBNAME-TEXTBOX
		$TTDBNameBox.SetToolTip($DBNameLabel, "Enter the main name of the databases`ne.q. Taxportal01; Taxportal02...TaxportalXY:`nTaxportal")
		
		$window.Refresh()
		
	   	}
		
	## IF THE DBCHECKBOX IS NOT ACTIVATED	
	Else { 
	
	## ALTERING THE DB-LABEL
	$DBHeadLine.Text = "Database:"
	
	$y = 80
	
		If ($multiInstCheck.Checked -eq $false){	
	 
	 	$y = $y - 80}
		
		Else { $y = 80}
		
		
		## SETTING THE HIGH OF ALL OBJECTS THAT ARE UNDER THE CHECKBOX LOCATED 
		$window.Size = New-Object System.Drawing.Size @(470,($windowY+$y))
		$DBheadline.Location = New-Object System.Drawing.Size(20,($dbheadlineY+$y))
		$multiDBCheck.SetBounds(170,($multiDBCheckY+$y),300,40)
		$DBQuantityLabel.Location = New-Object System.Drawing.Size(60,(240+$y))
		$DBQuantityBox.SetBounds(170,($DBQuantityY+$y),260,100)
		$DBNumLabel.Location = New-Object System.Drawing.Size(60,($dbnumY+$y))
		$DBNumBox.SetBounds(170,($dbnumY+$y),260,100)
		$DBNameLabel.Location = New-Object System.Drawing.Size(60,($dbnameY+$y))
		$DBNameBox.Visible = $true
		$DBNameBox2.Visible = $false
		$DBNameBox.SetBounds(170,($dbnameY+$y),260,100)
		$update.Location = New-Object System.Drawing.Size(20,($button+$y))
		$exit.Location = New-Object System.Drawing.Size(355,($button+$y))
		
		## ALTER A TOOLTIP FOR THE DBNAME-TEXTBOX
		$TTDBNameBox.SetToolTip($DBNameLabel, "Enter the Name of the instance of all databases")
		
		## DELETING THE NEW OBJECTS
		$window.Controls.Add($DBNameBox)
		$window.Controls.Add($DBNameBox2)
		$window.Controls.Remove($DBQuantityLabel)
		$window.Controls.Remove($DBQuantityBox)
		$window.Controls.Remove($DBNumLabel)
		$window.Controls.Remove($DBNumBox)
		$window.Refresh()
		}
})

#######################################################################

#### EXECUTION OF THE UPDATEBUTTON
$Update.add_Click({

	# PATH, WHERE THE SQL-UPDATE-SCRIPT(s) IS/ARE LOCATED 
	$updatepathDB = $updatepathDBBox.Text

	####################################################################################
	
	###INSTANCES:

	#######################ERRORS:
	
	## ERROR UPDATEPATH HAS NOT BEEN SET
	
	if ($updatepathDB -eq [String]::Empty){
	
	[System.Windows.Forms.MessageBox]::Show("UpdatePathError: `nPlease enter an updatepath!","TaxOne - Update-Script",
	[System.Windows.Forms.MessageBoxButtons]::OK,
	[System.Windows.Forms.MessageBoxIcon]::Hand)
	
	}	
	
	else
	{
		## ERROR UPDATEPATH DOES NOT EXIST
		### "-eq $false" MEANS THE SAMEAS "!" BUT ITS NOT COMPATIBLE WITH THE TEST-PATH CMD-LET 
		if (!(Test-Path $updatepathDB)) {
			
		[System.Windows.Forms.MessageBox]::Show("UpdatePathError: `nThe Path does not exists!","TaxOne - Update-Script",
		[System.Windows.Forms.MessageBoxButtons]::OK,
		[System.Windows.Forms.MessageBoxIcon]::Hand)
							
		}  
		
		else {
		
			## ERROR NO .SQL-FILES IN UPDATE-PATH
			### "-eq $false" MEANS THE SAMEAS "!" BUT ITS NOT COMPATIBLE WITH THE TEST-PATH CMD-LET 
			if (!(Test-Path $updatepathDB\*.sql)) {
				
			[System.Windows.Forms.MessageBox]::Show("UpdatePathError: `nThere is no .sql-File in the Directory!","TaxOne - Update-Script",
			[System.Windows.Forms.MessageBoxButtons]::OK,
			[System.Windows.Forms.MessageBoxIcon]::Hand)
								
			}  
			
			else {
			
				## ERROR INSTANCE NAME HAS NOT BEEN SET
				
				# NAME OF INSTANCES
				$instName = $instNameBox.Text
				
				if ($instName -eq [String]::Empty){
				
				[System.Windows.Forms.MessageBox]::Show("InstanceNameError: `nPlease enter an instancename!","TaxOne - Update-Script",
				[System.Windows.Forms.MessageBoxButtons]::OK,
				[System.Windows.Forms.MessageBoxIcon]::Hand)
				
				}
					
				else
				{
				
					## ERROR INSTANCECHECKBOX IS CHECKED AND FIRSTNUMBER HAS NOT BEEN SET
					# FIRST NUMBER OF INSTANCES
							$i = $instNumBox.Text
							
					if ($multiInstCheck.Checked -eq $true -and $i -eq ""){
					
					[System.Windows.Forms.MessageBox]::Show("MultipleInstanceError: `n Please enter a firstnumber of instances!","TaxOne - Update-Script",
					[System.Windows.Forms.MessageBoxButtons]::OK,
					[System.Windows.Forms.MessageBoxIcon]::Hand)
					}
			
					else
					{
						# QUANTITY OF INSTANCES
						$instQuantity = $instQuantityBox.Text
						$instQuantity = [Int]$instQuantity
			
						## ERROR IF QUANTITY IS SMALLER THAT 2
						# IF THERE IS MORE THAN ONE INSTANCE 
						if ($multiinstCheck.Checked -eq $true -and $instQuantity -lt 2){
						
						[System.Windows.Forms.MessageBox]::Show("MultipleInstanceError: `nPlease choose at min a quantity of 2 instances!","TaxOne - Update-Script",
						[System.Windows.Forms.MessageBoxButtons]::OK,
						[System.Windows.Forms.MessageBoxIcon]::Hand)
						
						}
						
						else{
								
							if ($multiinstCheck.Checked -eq $true -and $instQuantity -ge 2){

								## ALTERING THE DATATYPE OF i TO INT32
								$i = [Int]$i
								$ihilf = $i
								
								# ARRAY OF INSTANCES
								$instArray = @()

								# FILL ARRAY WITH INSTANCENAMES
								for ($i; $i -le ($ihilf + $instQuantity-1); $i++) {
								
									# ALL NAMES, WHICH ARE SMALLER THAN 10 WILL GET A "0" BETWEEN THE NAME AND NUMBER
							  	  	if ($i -lt 10) {
							   	    	$instArray += $instName +'0'+ [String]$i
							  	  	} 
									Else {
							      	 	$instArray += "$instName$i"
							   	 		}
								}
							}

							# ELSE THERE IS ONLY ONE INSTANCE
							Else {
								
								# INSTANCENAME
								$instName = $instNameBox.Text
								
								# ARRAY OF ONE INSTANCE
								$instArray += $instName
							}

							####################################################################################
							
							# DATABASES:
							
							## ERROR INSTANCE NAME HAS NOT BEEN SET
				
							# NAME OF DATABASES
							$DBName = $DBNameBox.Text
							
							if ($DBName -eq [String]::Empty){
							
							[System.Windows.Forms.MessageBox]::Show("DatabaseNameError: `nPlease enter a Databasename!","TaxOne - Update-Script",
							[System.Windows.Forms.MessageBoxButtons]::OK,
							[System.Windows.Forms.MessageBoxIcon]::Hand)
							
							}
								
							else
							{
								## ERROR DATABASECHECKBOX IS CHECKED AND FIRSTNUMBER HAS NOT BEEN SET

								# FIRST NUMBER OF DATABASES
								$z = $DBNumBox.Text
							
								if ($multiDBCheck.Checked -eq $true -and $z -eq ""){
						
								[System.Windows.Forms.MessageBox]::Show("MultipleDataBaseError: `n Please enter a firstnumber of databases!","TaxOne - Update-Script",
								[System.Windows.Forms.MessageBoxButtons]::OK,
								[System.Windows.Forms.MessageBoxIcon]::Hand)
								}	
									
								else
								{
								    # QUANTITY OF DATABASES
									$DBQuantity = $DBQuantityBox.Text
									$DBQuantity = [Int]$DBQuantity

									# IF THERE IS MORE THAN ONE DATABASE
									if ($multiDBCheck.Checked -eq $true -and $DBQuantity -lt 2){
									
									[System.Windows.Forms.MessageBox]::Show("MultipleDatabaseError: `nPlease choose at min a quantity of 2 databases!","TaxOne - Update-Script",
									[System.Windows.Forms.MessageBoxButtons]::OK,
									[System.Windows.Forms.MessageBoxIcon]::Hand)
									
									}
									
									else
									{
										
										if ($multiDBCheck.Checked -eq $true -and $DBQuantity -ge 2){

											# FIRST NUMBER OF DATABASES
											$z = [Int]$z
											$zhilf = $z
											
											# ARRAY OF DATABASES
											$DBArray = @()

											# FILL ARRAY WITH DATABASENAMES
											for ($z; $z -le ($zhilf + $DBQuantity-1); $z++) {
											
												# ALL NAMES, WHICH ARE SMALLER THAN 10 WILL GET A "0" BETWEEN THE NAME AND NUMBER
										  	  	if ($z -lt 10) {
										   	    	$DBArray += $DBName +'0'+ [String]$z
										  	  	} 
												Else {
										      	 	$DBArray += "$DBName$z"
										   	 		}
											}
										}

										# ELSE THERE IS ONLY ONE DATABASE
										Else {
																
											# ARRAY OF ONE DATABASE
											$DBArray += $DBName
										}

										##### EXECUTION OF THE UPDATE BUTTON

										# UPDATEFILES:
										
										# CREATE AN ARRAY OF .SQL-FILES LOCATED IN THE UPDATEDIRECTORY 
										$FileArray = Get-Childitem $updatepathDB | Where-Object {$_.Extension -eq ".sql"}
										
######################################## EXECUTION OF THE UPDATE BUTTON	
																								
										### LOGGING
																								
										## CHECKING IF LOG-DIRECTORY E:\TaxOne_Update\ EXISTS, ELSE CREATE IT									
										if (Test-Path "E:\TaxOne_Update") {
										} 
										else {
										New-Item -Path "E:\" -Name "TaxOne_Update" -ItemType directory
										}
										
										## CREATING A VARIABLE WITH THE DATE FOR THE LOGNAME
										$date = Get-Date -Format dd.MM.yyyy_HH-mm 
																																								
										## CREATING A LOG.TXT-FILE
										if (Test-Path "E:\TaxOne_Update\LOG_$date.txt") {
										}  
										else {
										New-Item -Path "E:\TaxOne_Update" -Name "LOG_$date.txt" -Itemtype "file" -value "Scripts will be executed on '$serverName'`n." –force
										
										}
																														
										# FOR EACH INSTANCE IN INSTANCEARRAY
										foreach ($instName in $instArray){

											# FOR EACH DATABASENAME IN DATABASEARRAY 
										  	foreach ($DBName in $DBArray){
											
												# FILL ARRAY WITH ALL FILES THAT WHERE FOUND
												for ($i = 0; $i -lt $FileArray.Count; $i++) {
											
													# WRITING THE FILENAME, DATABASENAME AND INSTANCENAME FOR THE LOGGING ON THE SCREEN
													Write-Host "Executing file: " $FileArray[$i] "on database: $DBName in instance: $instName"
													
													## WRITING FILEARRAY[i] IN $SCRIPT ELSE SQL-CMD AND LOGS WILL FEED THE WHOLE ARRAY
													$script =  $FileArray[$i]
													
													# WRITING THE FILENAME, DATABASENAME AND INSTANCENAME FOR THE LOGGING IN THE LOG-FILE
													Add-Content "E:\TaxOne_Update\LOG_$date.txt" "`nExecuting file:  $script on database: $DBName in instance: $instName"
																							
													# EXECUTING EVERY FILE ON EVERY DATABASE IN EVERY INSTANCE  || AND CREATING A LOGFILE NAMED AS THE SQL-FILE
													sqlcmd -S $serverName\$instName -d $DBName -E -i $updatepathDB\$script -o "E:\TaxOne_Update\$script.txt"
													
													## COPY THE DATA FROM $script.txt TO THE LOGFILE
													$datascripttxt = Get-Content E:\TaxOne_Update\$script.txt
													Add-Content "E:\TaxOne_Update\LOG_$date.txt" "`n$datascripttxt"
													
													## DELETING THE $script.txt
													remove-item "E:\TaxOne_Update\$script.txt"
											
													## WAIT 10 SECONDS UNTIL THE NEXT SCRIPT WILL RUN
													Start-Sleep -Seconds 1	
												}
											}
										}
										
											## END OF SCRIPT
											
											## ADDING AN END MESSAGE TO THE LOG-FILE
											
											Add-Content E:\TaxOne_Update\LOG_$date.txt "`nEnd of script" 
											
											[System.Windows.Forms.MessageBox]::Show("The Update has been finished!","TaxOne - Update-Script",
											[System.Windows.Forms.MessageBoxButtons]::OK,
											[System.Windows.Forms.MessageBoxIcon]::Asterisk)
											
									
									
									
									}# ERROR DATABASEQUANTITY
								}# ERROR DATABASE FIRST NUMBER	
							}# ERROR DATABASENAME
						}# ERROR INSTANCEQUANTITY
					}# ERROR FIRST INST NUMBER		
				}# ERROR INSTNAME
			}# ERROR NO .SQL FILES IN UPDATE-PATH
		}# ERROR WRONG UPDATEPATH
	}# ERROR NO UPDATEPATH	
})# CLICKEVENT

###END OF UPDATE 
#######################################################################

## CLICKEVENT FOR THE EXIT BUTTON - CLOSING THE WINDOW
$exit.Add_Click({$window.Close()})

## ADD THE OBJECTS TO THE WINDOW
#$window.Controls.Add($)

$window.Controls.Add($headline)

$window.Controls.Add($updatepathDBLabel)
$window.Controls.Add($updatepathDBBox)
$window.Controls.Add($browseUpdatePathbutton)
$window.Controls.Add($SetButton)

$window.Controls.Add($multiInst)
$window.Controls.Add($multiInstCheck)
$window.Controls.Add($instNameLabel)
$window.Controls.Add($instNameBox)

$window.Controls.Add($DBheadline)
$window.Controls.Add($multiDBCheck)
$window.Controls.Add($DBNameLabel)
$window.Controls.Add($DBNameBox)

$window.Controls.Add($update)
$window.Controls.Add($exit)

## show the window
$window.ShowDialog()