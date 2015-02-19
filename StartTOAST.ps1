
### UPDATESCRIPT FOR TAXONE SQL-DATABASES

## 19.01.2015 - Thomas Scholz & Robert Ostwald 

## Version: 1.0
# ChangeLog
# 	18.02.2015
#		- added "out-null" as pipe to assembly load to avoid the message about loading the .dll
#		- added click-event "next" logics
#			- after checking a combination of installation/update of webserver/database, a cmd will be started, which will start the right script with the ExecutionPolicy option
#


#########################################################################

#### GUI-CODE:

### GENERAL:
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

## FONTS
$headlineFont = New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Underline)
$font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::regular)
$underlined = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Underline)

## CREATING A NEW WINDOW 
$window1 = New-Object System.Windows.Forms.Form
$window1Y = 360
$window1.Size = New-Object System.Drawing.Size @(470,$window1Y)
$window1.Text = “TaxOne - Installation & Update”
$window1.StartPosition = “CenterScreen”
## STATIC WINDOWSIZE:
$window1.FormBorderStyle = 'FixedDialog'
## $window.BackgroundImage = $Image

## CREATE A HEADLINE
$headline1 = New-Object System.Windows.Forms.Label
$headline1.Text = “TaxOne - Installation and Update”
$headline1.AutoSize = $True
$headline1.Location = New-Object System.Drawing.Size(20,20)
$headline1.Font = $HeadlineFont

## CREATE A LABEL FOR THE QUESTION
$Question = New-Object System.Windows.Forms.Label
$Question.Text = "What do you want to do?"
$Question.AutoSize = $True
$Question.Location = New-Object System.Drawing.Size(20,80)
$Question.Font = $underlined

###CREATE A GROUPBOX FOR SERVERS
$groupBoxServer = New-Object System.Windows.Forms.GroupBox
$groupBoxServer.Location = New-Object System.Drawing.Size(22,120) 
$groupBoxServer.size = New-Object System.Drawing.Size(190,130) 
$groupBoxServer.text = "Server" 
$groupBoxServer.Font = $font

###CREATE A GROUPBOX FOR ACTION
$groupBoxAction = New-Object System.Windows.Forms.GroupBox
$groupBoxAction.Location = New-Object System.Drawing.Size(230,120) 
$groupBoxAction.size = New-Object System.Drawing.Size(200,130) 
$groupBoxAction.text = "Action" 
$groupBoxAction.Font = $font

## CREATE A RADIOBUTTON FOR THE DATABASE-SERVER
$DBRadio = New-Object System.Windows.Forms.RadioButton
$DBRadio.Text = "Database-Server"
$DBRadio.SetBounds(30,30,150,40)
$DBRadio.Checked = $false
$DBRadio.Font = $font

## CREATE A RADIOBUTTON FOR THE Websever
$WEBRadio = New-Object System.Windows.Forms.RadioButton
$WEBRadio.Text = "Webserver"
$WEBRadio.SetBounds(30,70,150,40)
$WEBRadio.Checked = $false
$WEBRadio.Font = $font

## CREATE A RADIOBUTTON FOR THE DATABASE-SERVER
$installRadio = New-Object System.Windows.Forms.RadioButton
$installRadio.Text = "Installation"
$installRadio.SetBounds(30,30,150,40)
$installRadio.Checked = $false
$installRadio.Font = $font

## CREATE A RADIOBUTTON FOR THE Websever
$updateRadio = New-Object System.Windows.Forms.RadioButton
$updateRadio.Text = "Update"
$updateRadio.SetBounds(30,70,150,40)
$updateRadio.Checked = $false
$updateRadio.Font = $font

# NEXTBUTTON
$next = New-Object System.Windows.Forms.Button
$button = 280
$next.Location = New-Object System.Drawing.Size(20,$button)
$next.Size = New-Object System.Drawing.Size(75,23)
$next.Text = “Next”
$next.Font = $Font

# EXITBUTTON
$exit1 = New-Object System.Windows.Forms.Button
$exit1.Location = New-Object System.Drawing.Size(355,$button)
$exit1.Size = New-Object System.Drawing.Size(75,23)
$exit1.Text = “Exit”
$exit1.Font = $Font

#######################################################################

### CLICK-EVENTS 

$next.Add_Click({


    ## INSTALL DB-SERVER
    IF ($DBRadio.Checked -eq $true -and $installRadio.Checked -eq $true) {
		Start-Process -FilePath "\\defr2app31\d$\TOAST\StartTOAST_Install.cmd"
    }
    
    ## UPDATE DB-SERVER
    IF ($DBRadio.Checked -eq $true -and $updateRadio.Checked -eq $true) {
		Start-Process -FilePath "\\defr2app31\d$\TOAST\StartTOAST_UpdateDB.cmd"
    }
    
    ## INSTALL WEB-SERVER
    IF ($WEBRadio.Checked -eq $true -and $installRadio.Checked -eq $true) {
		Start-Process -FilePath "\\defr2app31\d$\TOAST\StartTOAST_Install.cmd"
    }
    
    ## UPDATE WEB-SERVER
    IF ($WEBRadio.Checked -eq $true -and $updateRadio.Checked -eq $true){
		Start-Process -FilePath "\\defr2app31\d$\TOAST\StartTOAST_UpdateWeb.cmd"
    }
})

## CLICKEVENT FOR THE EXIT BUTTON - CLOSING THE WINDOW
$exit1.Add_Click({$window1.Close()})

#$window.Controls.Add($)
#$window.Controls.Add($)

$window1.Controls.Add($Question)
$window1.Controls.Add($groupBoxServer)
$window1.Controls.Add($groupBoxAction)
$groupBoxServer.Controls.Add($DBRadio)
$groupBoxServer.Controls.Add($WEBRadio)
$groupBoxAction.Controls.Add($installRadio)
$groupBoxAction.Controls.Add($updateRadio)
$window1.Controls.Add($Question)
$window1.Controls.Add($headline1)
$window1.Controls.Add($next)
$window1.Controls.Add($exit1)

## show the window
$window1.ShowDialog()
