# Get current user data
$scriptUser = $env:USERNAME
$scriptUserN = Get-ADUser $scriptUser
$scriptUserName = $scriptUserN.GivenName + " " + $scriptUserN.Surname + " - " + $scriptUser

#region: Basic validation to restrict access to specific users
$groups = "AD_Group_1", "AD_Group_2"

#$admUser
$group1Users = (Get-ADGroupMember -Recursive -Identity $groups[0]).Name
#$cybUser
$group2Users = (Get-ADGroupMember -Recursive -Identity $groups[1]).Name

if (($scriptUser -in $group1Users) -or ($scriptUser -in $group2Users)) {
    #endregion: Basic validation

    #region: Import SCCM Module
    Import-Module "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"

    if (($PWD.Path -ne "P01:\") -or ($null -eq $allDeviceCollection)) {
        # Select SITE Manually
        $site_server = "mysite.mydomain.com"
        Set-Location -Path "P01:\"

        # Select SITE Automatically
        #$SiteCode = Get-PSDrive -PSProvider CMSITE
        #Set-Location -Path "$($SiteCode.Name):\"
    
        Write-Host "`nFetching Device Collections..."
        $allDeviceCollection = (Get-CMDeviceCollection).Name
        Write-Host "Done"
        #Write-Host "`nFetching Software Update Groups..."
        #$allSUG = (Get-CMSoftwareUpdateGroup).LocalizedDisplayName | where {$_ -match 2023}
    }

    $listDeviceCollection = $allDeviceCollection | Sort-Object

    $global:timeS = 120 #120
    $global:timeD = 60 #60
    $global:countdownReboot = 600 #600
    $global:path = "C:\Scripts\patching_scripts\logs"
    $global:counter = 0

    #region: Get users to display in dropdown, optionally exclude certain users
    $finalUsers = @()
    $excludeUsers = "user1", "user2", $null
    $allUsers = Get-ADGroupMember -Identity "AD_Admin_Group" | Get-ADUser | Where-Object { $_.GivenName -notin $excludeUsers } | Sort-Object -Property GivenName
    foreach ($user in $allUsers) { $finalUsers += $user.givenname + " " + $user.Surname + " - " + $user.Name }
    #endregion: Get users

    <#
$dateTime = Get-Date -Format dd-MM-yyyy_HH-mm-ss
$date = Get-Date
#>

    $year = (Get-Date).Year
    $month = Get-Date -Format MMM
    $finalSug = "SUG_" + $month.ToUpper() + "_" + $year
    #$finalSug = ($allSUG | where {$_ -match $month})[0]
    #>
    # ================ FORM ===================

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Windows Servers Patching"
    #$form.Size = '1020,420' #width, height
    $form.WindowState = 'Maximized'
    $form.StartPosition = 'CenterScreen'
    $form.BackColor = "White"

    # Optional: Logos to display
    $image1 = [system.drawing.image]::fromfile("C:\Scripts\patching_scripts\images\logo1.png")
    $pictureBox1 = New-Object System.Windows.Forms.PictureBox
    $pictureBox1.Width = $image1.Size.Width
    $pictureBox1.Height = $image1.Size.Height
    $pictureBox1.Image = $image1
    $pictureBox1.Location = '50,10'
    #$pictureBox1.AutoSize = $true
    $form.Controls.Add($pictureBox1)

    $image2 = [system.drawing.image]::fromfile("C:\Scripts\patching_scripts\images\logo2.png")
    $pictureBox2 = New-Object System.Windows.Forms.PictureBox
    #$pictureBox2.Width = 40
    #$pictureBox2.Height = 40
    $pictureBox2.AutoSize = $true
    $pictureBox2.Image = $image2
    $pictureBox2.Location = '400,10'
    $form.Controls.Add($pictureBox2)

    # TextBox : For general information
    $infoTextBox = New-Object System.Windows.Forms.TextBox
    $infoTextBox.Location = '525,15'
    $infoTextBox.Size = '815,45'
    $infoTextBox.Text = 'Welcome!'
    $infoTextBox.Multiline = $true
    $infoTextBox.Enabled = $false
    $infoTextBox.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)
    $infoTextBox.BackColor = "white"
    $form.Controls.Add($infoTextBox)

    # RichTextBox: Main Output Box
    $outputBox2 = New-Object System.Windows.Forms.RichTextBox
    $outputBox2.Location = '525,75'
    $outputBox2.Size = '815,545'
    $outputBox2.Multiline = $true
    $outputBox2.ScrollBars = 'both'
    $outputBox2.WordWrap = $false
    $outputBox2.ReadOnly = $true
    #$outputBox2.Anchor = 'Left, Top, Right, Bottom'
    $outputBox2.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $outputBox2.BackColor = "white"
    $form.Controls.Add($outputBox2)

    $labelTimer = New-Object System.Windows.Forms.Label
    $labelTimer.Location = '1288,77'
    $labelTimer.Size = '50,40'
    #$labelTimer.Text = "10"
    $labelTimer.TextAlign = "MiddleCenter"
    $labelTimer.AutoSize = $false
    $labelTimer.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 25)
    #$labelTimer.BackColor = "Red"
    $form.Controls.Add($labelTimer)
    #$labelTimer.BringToFront()

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = '50,100'
    $tabControl.Size = '450,520'
    $tabControl.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 12)
    $tabControl.SelectedIndex = 0
    $form.controls.add($tabControl)

    $tabDeviceColl = New-Object System.Windows.Forms.TabPage
    $tabDeviceColl.Text = "Device Collection"
    $tabDeviceColl.BackColor = "white"
    $tabControl.controls.add($tabDeviceColl)

    $tabServers = New-Object System.Windows.Forms.TabPage
    $tabServers.Text = "Servers"
    $tabServers.BackColor = "white"
    #$tabServers.Add_Click({$form.add_shown({$textServers.select()})})
    $tabControl.controls.add($tabServers)

    # ================= DEVICE COLLECTION ===================

    # Label : New Device Collection
    $labelDcName = New-Object System.Windows.Forms.Label
    $labelDcName.AutoSize = $true
    $labelDcName.Location = '70,32'
    $labelDcName.Text = 'SCCM Device Collection'
    $labelDcName.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 18, [System.Drawing.FontStyle]::Bold)
    $tabDeviceColl.Controls.Add($labelDcName)

    # DROPDOWN : SELECT DC
    $dropDCList = New-Object System.Windows.Forms.Combobox
    $dropDCList.Location = '75,70'
    $dropDCList.Size = '250,20'

    # Populate DCs
    foreach ($dc in $listDeviceCollection) { [void] $dropDCList.Items.Add($dc) }
    $dropDCList.Add_Click({ $infoTextBox.Text = "Enter a device collection or select from the dropdown" })#; if($dropDCList.Text -match "SAP"){$radioReboot.Enabled = $false}})

    # Optional: Disable Reboot if DC contains certain words
    $excludeDeviceColl = "DONOREBOOT", "DC_EXCLUDE" -join "|"
    $dropDCList.Add_TextChanged({ if ($dropDCList.Text -match $excludeDeviceColl) { $radioReboot.Enabled = $false } else { $radioReboot.Enabled = $true } })

    $dropDCList.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    #$form.Controls.Add($dropDCList)
    $tabDeviceColl.Controls.Add($dropDCList)

    $imageRefresh = [system.drawing.image]::fromfile("C:\Scripts\patching_scripts\images\reload.png")
    #$imageRefresh.SetResolution('20,20')

    # BUTTON : REFRESH
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = '330,68'
    $refreshButton.Size = '27,27'
    $refreshButton.BackgroundImage = $imageRefresh
    $refreshButton.BackgroundImageLayout = "zoom"
    #$refreshButton.Add_MouseHover({"Refresh"})
    $refreshButton.Add_Click({ $dropDCList.Text = "refreshing..."
            $infoTextBox.Text = "Refresh Device Collection"
            $dropDCList.Items.Clear()
            $listDeviceCollection = (Get-CMDeviceCollection).Name | Where-Object { $_ -match "slot" } | Sort-Object
            foreach ($dc in $listDeviceCollection) { [void] $dropDCList.Items.Add($dc) }
            $dropDCList.Text = $null })
    #$form.Controls.Add($refreshButton)
    $tabDeviceColl.Controls.Add($refreshButton)

    # Label : INFO - Type or Select Device Collection
    $labelDcInfo = New-Object System.Windows.Forms.Label
    $labelDcInfo.AutoSize = $true
    $labelDcInfo.Location = '115,100'
    $labelDcInfo.Text = '*type or select device collection'
    $labelDcInfo.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    #$form.Controls.Add($labelDcInfo)
    $tabDeviceColl.Controls.Add($labelDcInfo)

    # RadioButton : Display members
    $radioMembers = New-Object System.Windows.Forms.RadioButton
    $radioMembers.AutoSize = $true
    $radioMembers.Location = '65,150'
    $radioMembers.Text = 'Display members'
    $radioMembers.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11)
    $radioMembers.Add_Click({ $infoTextBox.Text = "All servers in this device collection will be displayed" })
    #$form.Controls.Add($radioMembers)
    $tabDeviceColl.Controls.Add($radioMembers)

    # RadioButton : Compliance Cycles
    $radioCompCycles = New-Object System.Windows.Forms.RadioButton
    $radioCompCycles.AutoSize = $true
    $radioCompCycles.Location = '65,190'
    $radioCompCycles.Text = 'Compliance Cycles'
    $radioCompCycles.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11)
    $radioCompCycles.Add_Click({ $infoTextBox.Text = "Force the server to check for updates" })
    #$form.Controls.Add($radioCompCycles)
    $tabDeviceColl.Controls.Add($radioCompCycles)

    # RadioButton : Compliance Check
    $radioCompCheck = New-Object System.Windows.Forms.RadioButton
    $radioCompCheck.AutoSize = $true
    $radioCompCheck.Location = '65,230'
    $radioCompCheck.Text = 'Compliance Check'
    $radioCompCheck.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11)
    $radioCompCheck.Add_Click({ $infoTextBox.Text = "Check whether patches are installed on all the servers `r`nIncludes compliance cycles & deployment summarization" })
    #$form.Controls.Add($radioCompCheck)
    $tabDeviceColl.Controls.Add($radioCompCheck)

    <#
# CheckBox : Quick Check
$checkQuickCompCheck = New-Object System.Windows.Forms.CheckBox
$checkQuickCompCheck.AutoSize = $true
$checkQuickCompCheck.Location = '65,205'
$checkQuickCompCheck.Text = 'Quick Compliance Check'
$checkQuickCompCheck.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8)
$checkQuickCompCheck.Add_Click({
    $infoTextBox.Text = "Check compliance status without running cycles or deployment summarization"
    $radioCompCheck.Checked = $true
    $textSugName.Text = $null
})
$form.Controls.Add($checkQuickCompCheck)
#>

    # RadioButton : Apply patches
    $radioPatching = New-Object System.Windows.Forms.RadioButton
    $radioPatching.AutoSize = $true
    $radioPatching.Location = '240,150' #'225,107'
    $radioPatching.Text = 'Apply patches'
    $radioPatching.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11)
    $radioPatching.Add_Click({ $infoTextBox.Text = "Patches from the mentioned SUG will be pushed" })
    #$form.Controls.Add($radioPatching)
    $tabDeviceColl.Controls.Add($radioPatching)

    # TextBox : SUG Name
    $textSugName = New-Object System.Windows.Forms.TextBox
    #$textSugName.Text = 'sug_name'
    $textSugName.Location = '240,190' #'240,137'
    $textSugName.Width = '120'
    $textSugName.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $textSugName.Add_Click({ $infoTextBox.Text = "Make sure to enter a valid Software Update Group (SUG)" })
    #$form.Controls.Add($textSugName)
    $tabDeviceColl.Controls.Add($textSugName)

    # RadioButton : Reboot servers
    $radioReboot = New-Object System.Windows.Forms.RadioButton
    $radioReboot.AutoSize = $true
    $radioReboot.Location = '240,230' #'225,177'
    $radioReboot.Text = 'Reboot servers'
    $radioReboot.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11)
    $radioReboot.Add_Click({ $infoTextBox.Text = "All servers will be rebooted if needed" })
    #$form.Controls.Add($radioReboot)
    $tabDeviceColl.Controls.Add($radioReboot)

    # Label : Maker
    $labelMaker = New-Object System.Windows.Forms.Label
    $labelMaker.AutoSize = $true
    $labelMaker.Location = '50,300'
    $labelMaker.Text = 'Maker'
    $labelMaker.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11, [System.Drawing.FontStyle]::Bold)
    #$form.Controls.Add($labelMaker)
    $tabDeviceColl.Controls.Add($labelMaker)

    # Label : Checker
    $labelChecker = New-Object System.Windows.Forms.Label
    $labelChecker.AutoSize = $true
    $labelChecker.Location = '50,340'
    $labelChecker.Text = 'Checker'
    $labelChecker.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11, [System.Drawing.FontStyle]::Bold)
    #$form.Controls.Add($labelChecker)
    $tabDeviceColl.Controls.Add($labelChecker)

    #$makerName = Get-ADUser -Identity $env:USERNAME
    #$makerFinalName = $makerName.GivenName + " " + $makerName.Surname + " - " + $env:USERNAME

    # DROPDOWN : SELECT MAKER
    $dropMaker = New-Object System.Windows.Forms.Combobox
    $dropMaker.Enabled = $false
    $dropMaker.Location = '130,300'
    $dropMaker.Font = [System.Drawing.Font]::new("Georgia", 8)
    $dropMaker.Size = '230,55'
    foreach ($usr in $finalUsers) { [void] $dropMaker.Items.Add($usr) }
    #if($dropMaker.Enabled -eq $true){$dropMaker.Text = $dropChecker.Items | Where-Object {$_ -match $scriptUser}}
    $dropMaker.Add_Click({ $infoTextBox.Text = "Select your name from the dropdown" })
    #$form.Controls.Add($dropMaker)
    $tabDeviceColl.Controls.Add($dropMaker)

    # DROPDOWN : SELECT CHECKER
    $dropChecker = New-Object System.Windows.Forms.Combobox
    $dropChecker.Enabled = $false
    $dropChecker.Location = '130,340'
    $dropChecker.Size = '230,15'
    $dropChecker.Font = [System.Drawing.Font]::new("Georgia", 8)
    foreach ($usr in $finalUsers) { [void] $dropChecker.Items.Add($usr) }
    $dropChecker.Add_Click({ $infoTextBox.Text = "Select the checker's name from the dropdown" })
    #$form.Controls.Add($dropChecker)
    $tabDeviceColl.Controls.Add($dropChecker)

    # OK/Submit button
    $okButtonDC = New-Object System.Windows.Forms.Button
    $okButtonDC.Location = '110,405'
    $okButtonDC.Size = '85,27'
    $okButtonDC.Text = 'OK'
    $okButtonDC.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11, [System.Drawing.FontStyle]::Bold)
    #$okButtonDC.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #$form.AcceptButton = $okButtonDC
    $form.Controls.Add($okButtonDC)
    $tabDeviceColl.Controls.Add($okButtonDC)

    # Cancel button
    $cancelButtonDC = New-Object System.Windows.Forms.Button
    $cancelButtonDC.Location = '220,405'
    $cancelButtonDC.Size = '85,27'
    $cancelButtonDC.Text = 'Cancel'
    $cancelButtonDC.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11, [System.Drawing.FontStyle]::Bold)
    $cancelButtonDC.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButtonDC
    $form.Controls.Add($cancelButtonDC)
    $tabDeviceColl.Controls.Add($cancelButtonDC)

    #================= Individual SERVERS ===================

    # Label : Individual Servers
    $labelAppServers = New-Object System.Windows.Forms.Label
    $labelAppServers.AutoSize = $true
    $labelAppServers.Location = '105,20'
    $labelAppServers.Text = '   Individual Servers'
    $labelAppServers.Font = [System.Drawing.Font]::new('Microsoft Sans Serif', 16, [System.Drawing.FontStyle]::Bold)
    $tabServers.Controls.Add($labelAppServers)

    # TextBox : Servers List
    $textServers = New-Object System.Windows.Forms.TextBox
    $textServers.Location = '95,60'
    $textServers.Size = '250,180'
    $textServers.Multiline = $true
    $textServers.ScrollBars = 'vertical'
    $textServers.Add_Click({ $infoTextBox.Text = "Only use for Individual servers" })
    $textServers.Add_Keydown({ if (($_.Control) -and ($_.KeyCode -eq 'A')) { $textServers.SelectAll() } })
    $tabServers.Controls.Add($textServers)

    # RadioButton : Compliance Cycles for Individual Servers
    $radioCompCycServers = New-Object System.Windows.Forms.RadioButton
    $radioCompCycServers.Location = '40,270'
    $radioCompCycServers.AutoSize = $true
    $radioCompCycServers.Text = 'Compliance Cycles'
    $radioCompCycServers.Font = [System.Drawing.Font]::new('Microsoft Sans Serif', 13)
    $radioCompCycServers.Add_Click({
            $infoTextBox.Text = "Force the server to check for updates"
            $dropDCCompliance.Text = $null
            $dropDCCompliance.Enabled = $false
        })
    $tabServers.Controls.Add($radioCompCycServers)

    # RadioButton : Compliance Check for Individual Servers
    $radioCompCheckServers = New-Object System.Windows.Forms.RadioButton
    $radioCompCheckServers.Location = '235,270'
    $radioCompCheckServers.AutoSize = $true
    $radioCompCheckServers.Text = 'Compliance Check'
    $radioCompCheckServers.Font = [System.Drawing.Font]::new('Microsoft Sans Serif', 13)
    $radioCompCheckServers.Add_Click({
            $infoTextBox.Text = "Check whether patches are installed on all the servers `r`nIncludes compliance cycles & deployment summarization"
            $dropDCCompliance.Enabled = $true
        })
    $tabServers.Controls.Add($radioCompCheckServers)

    $labelDCCompliance = New-Object System.Windows.Forms.Label
    $labelDCCompliance.Text = "Device Collection`r`n      (optional)"
    $labelDCCompliance.AutoSize = $true
    $labelDCCompliance.Location = '40,320'
    $labelDCCompliance.Font = [System.Drawing.Font]::new('Microsoft Sans Serif', 13)
    $tabServers.Controls.Add($labelDCCompliance)

    $dropDCCompliance = New-Object System.Windows.Forms.Combobox
    $dropDCCompliance.Location = '215,333'
    $dropDCCompliance.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $dropDCCompliance.Width = '190'

    foreach ($dcs in $listDeviceCollection) { [void] $dropDCCompliance.Items.Add($dcs) }
    $dropDCCompliance.Add_Click({ $infoTextBox.Text = "Compliance will be checked for this device collection" })
    $dropDCCompliance.Add_TextChanged({ $radioCompCheckServers.Checked = $true })


    # Disable Reboot if DC contains certain words
    #$excludeDeviceColl = "DONOREBOOT", "DC_EXCLUDE" -join "|"
    #$dropDCList.Add_TextChanged({if($dropDCList.Text -match $excludeDeviceColl){$radioReboot.Enabled = $false} else{$radioReboot.Enabled = $true}})
    #$dropDCList.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)

    $tabServers.Controls.Add($dropDCCompliance)

    # Ok/Submit Button
    $okButtonServer = New-Object System.Windows.Forms.Button
    $okButtonServer.Location = '110,415'
    $okButtonServer.Size = '85,27'
    $okButtonServer.Text = 'OK'
    $okButtonServer.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11, [System.Drawing.FontStyle]::Bold)
    #$okButtonServer.DialogResult = [System.Windows.Forms.DialogResult]::OK
    #$form.AcceptButton = $okButtonServer
    $form.Controls.Add($okButtonServer)
    $tabServers.Controls.Add($okButtonServer)

    # Cancel button
    $cancelButtonServer = New-Object System.Windows.Forms.Button
    $cancelButtonServer.Location = '220,415'
    $cancelButtonServer.Size = '85,27'
    $cancelButtonServer.Text = 'Cancel'
    $cancelButtonServer.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 11, [System.Drawing.FontStyle]::Bold)
    $cancelButtonServer.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $cancelButtonServer
    $form.Controls.Add($cancelButtonServer)
    $tabServers.Controls.Add($cancelButtonServer)

    #================= FORM END ===================

    #$form.TopMost = $true

    # Update description box and enable/disable buttons

    $radioPatching.Add_Click({ 
            $dropMaker.Enabled = $true; $dropChecker.Enabled = $true; $textSugName.Text = $finalSug; 
            $dropMaker.Text = $dropMaker.Items | Where-Object { $_ -match $scriptUser } 
        })

    $radioReboot.Add_Click({ 
            $dropMaker.Enabled = $true; $dropChecker.Enabled = $true; $textSugName.Text = $null; 
            $dropMaker.Text = $dropMaker.Items | Where-Object { $_ -match $scriptUser } 
        })

    $radioMembers.Add_Click({ 
            $dropMaker.Enabled = $false; $dropChecker.Enabled = $false; $textSugName.Text = $null; 
            $dropMaker.Text = $null; $dropChecker.Text = $null 
        })

    $radioCompCycles.Add_Click({
            $dropMaker.Enabled = $false; $dropChecker.Enabled = $false; $textSugName.Text = $null; 
            $dropMaker.Text = $null; $dropChecker.Text = $null 
        })

    $radioCompCheck.Add_Click({ 
            $dropMaker.Enabled = $false; $dropChecker.Enabled = $false; $textSugName.Text = $null; 
            $dropMaker.Text = $null; $dropChecker.Text = $null
        })

    <#
$checkQuickCompCheck.Add_Click({
    $dropMaker.Enabled = $false; $dropChecker.Enabled = $false; $textSugName.Text = $null; 
    $dropMaker.Text = $null; $dropChecker.Text = $null
})
#>

    #if(($status -eq "ok"))

    # Call this function to output text to main output box
    Function Add-OutputBoxLine {
        Param ($Message, $scrollCheck, $fontSize2, $lucidaBool)
        if ($lucidaBool) { $outputBox2.SelectionFont = [System.Drawing.Font]::new("Lucida Console", $fontSize2) }
        if ($lucidaBool -eq $false) { $outputBox2.SelectionFont = [System.Drawing.Font]::new("Microsoft Sans Serif", $fontSize2) }
        $outputBox2.SelectionAlignment = "Left"
        $outputBox2.AppendText("$Message")
        $outputBox2.Refresh()
        if ($scrollCheck -eq $true) { $outputBox2.ScrollToCaret() }
    
    }

    # Call this function to output a colored line to main output box
    Function Append-ColoredLine {
        param( 
            [Parameter(Mandatory = $true, Position = 0)]
            [System.Windows.Forms.RichTextBox]$box,
            [Parameter(Mandatory = $true, Position = 1)]
            [System.Drawing.Color]$color,
            [Parameter(Mandatory = $true, Position = 2)]
            [string]$text,
            [Parameter(Mandatory = $true, Position = 3)]
            [bool]$bold,
            [Parameter(Mandatory = $true, Position = 4)]
            [int]$fontSize,
            [Parameter(Mandatory = $true, Position = 5)]
            [string]$textAlign
        )
        $box.Refresh()
        $box.SelectionStart = $box.TextLength
        $box.SelectionLength = 0
    
        if ($bold) { $box.SelectionFont = [System.Drawing.Font]::new("Microsoft Sans Serif", $fontSize, [System.Drawing.FontStyle]::Bold) }
        if ($bold -eq $false) { $box.SelectionFont = [System.Drawing.Font]::new("Microsoft Sans Serif", $fontSize) }

        $box.SelectionAlignment = $textAlign
        $box.SelectionColor = $color
        $box.AppendText($text)
        #$box.AppendText([Environment]::NewLine)
        $box.Refresh()
        $box.ScrollToCaret()
    }

    # Call this function to output an error to main output box
    Function Add-OutputBoxErrorLine {
        Param ($MessageError)
        $outputBox2.Refresh()
        $outputBox2.Clear()
        $outputBox2.SelectionAlignment = 'center'
        $outputBox2.SelectionColor = 'Red'
        $outputBox2.SelectionFont = [System.Drawing.Font]::new("Microsoft Sans Serif", 12, [System.Drawing.FontStyle]::Bold)
        $outputBox2.AppendText("$MessageError")
        $outputBox2.ScrollToCaret()
    }

    # Function : Compliance Check
    # Credits : https://thesysadminchannel.com/get-sccm-software-update-status-powershell
    function Get-SCCMSoftwareUpdateStatus {
                    
        [CmdletBinding()]
                    
        param(
            [Parameter()]
            [switch]  $DeploymentIDFromGUI,
 
            [Parameter(Mandatory = $false)]
            [Alias('ID', 'AssignmentID')]
            [string]   $DeploymentID,

            [Parameter(Mandatory = $false)]
            [Alias('Collection')]
            [string]   $CollectionName,

            [Parameter(Mandatory = $false)]
            [Alias('ServerName')]
            [string]   $Server,

            [Parameter(Mandatory = $false)]
            [Alias('checkValidDeployments')]
            [bool]   $checkValidDeployment,
         
            [Parameter(Mandatory = $false)]
            [ValidateSet('Compliant', 'InProgress', 'Error', 'Unknown')]
            [Alias('Filter')]
            [string]  $Status,

            $invalidDep = @()
 
        )
 
        BEGIN {
            $Site_Code = 'P01'
            $Site_Server = ''
            $HasErrors = $False
 
            if ($Status -eq 'Compliant') {
                $StatusType = 1
            }
 
            if ($Status -eq 'InProgress') {
                $StatusType = 2
            }
 
            if ($Status -eq 'Unknown') {
                $StatusType = 4
            }
 
            if ($Status -eq 'Error') {
                $StatusType = 5
            }
 
        }
 
        PROCESS {
            try {
                if ($DeploymentID -and $DeploymentIDFromGUI) {
                    Write-Error "Select the DeploymentIDFromGUI or DeploymentID Parameter. Not Both"
                    $HasErrors = $True
                    throw
                }
 
                if ($DeploymentIDFromGUI) {
                    $ShellLocation = Get-Location
                    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
                 
                    #Checking to see if module has been imported. If not abort.
                    if (Get-Module ConfigurationManager) {
                        Set-Location "$($Site_Code):\"
                        $DeploymentID = Get-CMSoftwareUpdateDeployment | Select-Object AssignmentID, AssignmentName | Out-GridView -OutputMode Single -Title "Select a Deployment and Click OK" | Select -ExpandProperty AssignmentID
                        Set-Location $ShellLocation
                    }
                    else {
                        Write-Error "The SCCM Module wasn't imported successfully. Aborting."
                        $HasErrors = $True
                        throw
                    }
                }

                if ($CollectionName) {

                    $ShellLocation = Get-Location
                    Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)
                 
                    #Checking to see if module has been imported. If not abort.
                    if (Get-Module ConfigurationManager) {
                        Set-Location "$($Site_Code):\"
                        #$DeploymentID = Get-CMSoftwareUpdateDeployment | select AssignmentID, AssignmentName | Out-GridView -OutputMode Single -Title "Select a Deployment and Click OK" | Select -ExpandProperty AssignmentID
                        Set-Location $ShellLocation
                    }
                    else {
                        Write-Error "The SCCM Module wasn't imported successfully. Aborting."
                        $HasErrors = $True
                        throw
                    }
                    #Write-host $collectionname

                    #Find latest deployment ID on a collection
                    $DeploymentID = (Get-CMSoftwareUpdateDeployment -CollectionName $CollectionName | Sort-Object StartTime -Descending).AssignmentID[0]
                    #$DeploymentID
                }
                else {
                    #Write-Error "A Collection Name was not specified.Aborting"
                    $HasErrors = $True
                    #throw  
                }

            
                if ($DeploymentID) {
                    $DeploymentNameWithID = Get-ciminstance -ComputerName $Site_Server -Namespace root\sms\site_$Site_Code -class SMS_SUMDeploymentAssetDetails -Filter "AssignmentID = $DeploymentID" | select AssignmentID, AssignmentName
                    $DeploymentName = $DeploymentNameWithID.AssignmentName | Select-Object -Unique
                }
                else {
                    Write-Error "A Deployment ID was not specified. Aborting."
                    $HasErrors = $True
                    throw  
                }
 
                if ($Status) {
                    $Output = Get-ciminstance -ComputerName $Site_Server -Namespace root\sms\site_$Site_Code -class SMS_SUMDeploymentAssetDetails -Filter "AssignmentID = $DeploymentID and StatusType = $StatusType" | `
                        <#Where-Object {$_.DeviceName -eq "$server"} |#> Select-Object DeviceName, CollectionName, StatusTime, @{Name = 'Status' ; Expression = { if ($_.StatusType -eq 1) { 'Compliant' } elseif ($_.StatusType -eq 2) { 'InProgress' } elseif ($_.StatusType -eq 5) { 'Error' } elseif ($_.StatusType -eq 4) { 'Unknown' } } }
 
                }
                else {      
                    $Output = Get-ciminstance -ComputerName $Site_Server -Namespace root\sms\site_$Site_Code -class SMS_SUMDeploymentAssetDetails -Filter "AssignmentID = $DeploymentID" | `
                        Where-Object { $_.DeviceName -eq "$server" } | Select-Object DeviceName, CollectionName, StatusTime, @{Name = 'Status' ; Expression = { if ($_.StatusType -eq 1) { 'Compliant' } elseif ($_.StatusType -eq 2) { 'InProgress' } elseif ($_.StatusType -eq 5) { 'Error' } elseif ($_.StatusType -eq 4) { 'Unknown' } } }
                    #Write-Host $site_server
                    #$server
                    #$Output
                }
 
                if ((-not $Output) -and ($checkValidDeployment)) {
                    #Write-Error "A Deployment with ID: $($DeploymentID) is not valid. Aborting"
                    #Write-Host "Deployment $server not valid" -ForegroundColor Red
                    #Append-ColoredLine -box $outputBox2 -color Red -text "Deployment $server not valid`r" -bold $false -fontSize 10
                    Append-ColoredLine -box $outputBox2 -color Red -text "Deployment $server not valid`r" -bold $false -fontSize 11 -textAlign Left

                    $invalidDep += $server
                    $HasErrors = $True
                                
                    throw
                 
                }

            }
            catch {
               
            }
            finally {
                if (($HasErrors -eq $false) -and ($Output)) {
                    #Write-Output "Deployment Name: $DeploymentName"
                    #Write-Output "Deployment ID:   $DeploymentID"
                    #Write-Output $Output #| Sort-Object Status
                    #Write-host $Output
                    #Add-OutputBoxLine -Message ($Output | Out-String) -scrollCheck $true
                    #Append-ColoredLine -box $outputBox2 -color Black -text ($Output | Out-String) -bold $false -fontSize 10
                }
                            
            }
            return $invalidDep, $Output
        }
        END {}
 
    }

    # Function : Countdown
    function Start-Countdown { 
        Param(
            [Int32]$Seconds = 10,
            [string]$Message = "Pausing for 10 seconds..."
        )
        ForEach ($Count in (1..$Seconds)) {
            Write-Progress -Id 1 -Activity $Message -Status "Waiting for $Seconds seconds, $($Seconds - $Count) left" -PercentComplete (($Count / $Seconds) * 100)
            Start-Sleep -Seconds 1
        }
        Write-Progress -Id 1 -Activity $Message -Status "Completed" -PercentComplete 100 -Completed
    }


    #region : Jobs

    # Function : Run Cycles
    function Run-Cycles ($serversList, $errorDisplay) {
        $cycleSuccess = @()
        $cycleFailed = @()

        foreach ($mem in $serversList) {
            try {
                Invoke-Command -ComputerName $mem -ScriptBlock {
                    Invoke-CimMethod -Namespace 'root\CCM' -ClassName SMS_Client -MethodName TriggerSchedule -Arguments @{sScheduleID = '{00000000-0000-0000-0000-000000000108}' } | Out-Null #SilentlyContinue 
                    Invoke-CimMethod -Namespace 'root\CCM' -ClassName SMS_Client -MethodName TriggerSchedule -Arguments @{sScheduleID = '{00000000-0000-0000-0000-000000000113}' } | Out-Null                           
                } -ErrorAction Stop
                $cycleSuccess += $mem
            }
            catch {
                $cycleFailed += $mem
                if ($errorDisplay -eq $true) {
                    Append-ColoredLine -box $outputBox2 -color Red -text "$mem - failed" -bold $false -fontSize 10
                }
            }
        }
        return $cycleSuccess, $cycleFailed
    }

    Function devicecoll-code {
        #$form.Refresh()
        $result = Switch ($null, "", " ", $false, "SUG_Name") {
            { $_ -contains $dropDCList.Text } { "Device Collection cannot be empty"; Break }
            { ($_ -contains $radioMembers.Checked) -and ($_ -contains $radioPatching.Checked) -and ($_ -contains $radioReboot.Checked) -and ($_ -contains $radioCompCycles.Checked) -and ($_ -contains $radioCompCheck.Checked) } { "Select something"; Break } # If All False
            { ($radioPatching.Checked) -and ($_ -contains $textSugName.Text) } { "Enter a valid SUG Name"; Break }
            { ($radioMembers.Checked -eq $false) -and ($radioCompCycles.Checked -eq $false) -and ($radioCompCheck.Checked -eq $false) -and (($_ -contains $dropMaker.Text) -or ($_ -contains $dropChecker.Text)) } { "Maker/Checker cannot be empty"; Break }
        
            #{($dropMaker.Text -ne $null) -and ($dropChecker.Text -ne $null) -and ($dropMaker.Text -eq $dropChecker.Text)} {"Maker and Checker cannot be same"; Break}
            #{($_ -notcontains $dropMaker.Text) -and ($_ -notcontains $dropChecker.Text) -and ($dropMaker.Text -eq $dropChecker.Text)} {"Maker and Checker cannot be same"; Break}
        }

        # If error, print it
        if ($null -ne $result) { Add-OutputBoxErrorLine -MessageError $result }

        if (($null -eq $result) -and ($dropMaker.Text -eq $dropChecker.Text) -and ($radioMembers.Checked -eq $false) -and ($radioCompCycles.Checked -eq $false) -and ($radioCompCheck.Checked -eq $false)) {
            $result2 = "False"
            Add-OutputBoxErrorLine -MessageError "Maker and Checker cannot be same"
        }
        else { $result2 = "True" }

        if ((($null -eq $result) -and ($result2 -eq "True")) -and ($radioMembers.Checked -eq $false) -and ($radioCompCycles.Checked -eq $false) -and ($radioCompCheck.Checked -eq $false) -and (($dropMaker.Text -notmatch "ADM_") -or ($dropChecker.Text -notmatch "ADM_"))) {
            $result3 = "False"
            Add-OutputBoxErrorLine -MessageError "Maker/Checker should contain your ADM ID"
        }
        else { $result3 = "True" }

        #============ Error handling complete ===============
    
        #if(![string]::IsNullOrWhiteSpace($deviceColl))
        if (($null -eq $result) -and ($result2 -eq "True") -and ($result3 -eq "True")) {
            $global:deviceColl = $dropDCList.Text.Trim().ToUpper()

            # Check if DC exists
            $dcExists = Get-CMDeviceCollection -Name $deviceColl | Sort-Object

            if ($null -ne $dcExists) {

                # Button : Yes
                $global:yesButtonEmptyPatch = New-Object System.Windows.Forms.Button
                $yesButtonEmptyPatch.Location = '820,635'
                $yesButtonEmptyPatch.Size = '85,27'
                $yesButtonEmptyPatch.Text = 'Yes'
                $yesButtonEmptyPatch.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 13)
                $yesButtonEmptyPatch.Visible = $true
                $form.Controls.Add($yesButtonEmptyPatch)

                # Button : No
                $global:noButtonEmptyPatch = New-Object System.Windows.Forms.Button
                $noButtonEmptyPatch.Location = '930,635'
                $noButtonEmptyPatch.Size = '85,27'
                $noButtonEmptyPatch.Text = 'No'
                $noButtonEmptyPatch.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 13)
                $noButtonEmptyPatch.Visible = $true
                $form.Controls.Add($noButtonEmptyPatch)

                # Button : Yes
                $global:yesButtonFullPatch = New-Object System.Windows.Forms.Button
                $yesButtonFullPatch.Location = '820,635'
                $yesButtonFullPatch.Size = '85,27'
                $yesButtonFullPatch.Text = 'Yes'
                $yesButtonFullPatch.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 13)
                $yesButtonFullPatch.Visible = $true
                $form.Controls.Add($yesButtonFullPatch)

                # Button : No
                $global:noButtonFullPatch = New-Object System.Windows.Forms.Button
                $noButtonFullPatch.Location = '930,635'
                $noButtonFullPatch.Size = '85,27'
                $noButtonFullPatch.Text = 'No'
                $noButtonFullPatch.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 13)
                $noButtonFullPatch.Visible = $true
                $form.Controls.Add($noButtonFullPatch)

                # Button : Yes
                $global:yesButtonReboot = New-Object System.Windows.Forms.Button
                $yesButtonReboot.Location = '820,635'
                $yesButtonReboot.Size = '85,27'
                $yesButtonReboot.Text = 'Yes'
                $yesButtonReboot.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 13)
                $yesButtonReboot.Visible = $true
                $form.Controls.Add($yesButtonReboot)

                # Button : No
                $global:noButtonReboot = New-Object System.Windows.Forms.Button
                $noButtonReboot.Location = '930,635'
                $noButtonReboot.Size = '85,27'
                $noButtonReboot.Text = 'No'
                $noButtonReboot.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 13)
                $noButtonReboot.Visible = $true
                $form.Controls.Add($noButtonReboot)


                <#
            $yesButtonEmptyPatch.Visible = $false
            $yesButtonFullPatch.Visible = $false
            $yesButtonReboot.Visible = $false
            $noButtonEmptyPatch.Visible = $false
            $noButtonFullPatch.Visible = $false
            $noButtonReboot.Visible = $false
            #>

                # Display Members
                if ($radioMembers.Checked) {
                    $yesButtonEmptyPatch.Visible = $false
                    $yesButtonFullPatch.Visible = $false
                    $yesButtonReboot.Visible = $false
                    $noButtonEmptyPatch.Visible = $false
                    $noButtonFullPatch.Visible = $false
                    $noButtonReboot.Visible = $false
                    Append-ColoredLine -box $outputBox2 -color DarkBlue -text "$deviceColl`r`n" -bold $true -fontSize 15 -textAlign Center
                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nDisplay Members:`r`n" -bold $true -fontSize 13 -textAlign Left

                    $members = (Get-CMCollectionMember -CollectionName $deviceColl).Name | Sort-Object
                
                    if ([string]::IsNullOrWhiteSpace($members)) { 
                        Append-ColoredLine -box $outputBox2 -color Red -text "`r`nEmpty Collection" -bold $true -fontSize 11 -textAlign Left 
                    }
                    else {
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nCount = $($members.Count)`r`n`n" -bold $true -fontSize 11 -textAlign Left
                        Add-OutputBoxLine -Message ($members | Out-String) -scrollCheck $false -fontSize2 12 -lucidaBool $true
                    }
                }

                # Compliance Cycles
                if ($radioCompCycles.Checked) {
                    $yesButtonEmptyPatch.Visible = $false
                    $yesButtonFullPatch.Visible = $false
                    $yesButtonReboot.Visible = $false
                    $noButtonEmptyPatch.Visible = $false
                    $noButtonFullPatch.Visible = $false
                    $noButtonReboot.Visible = $false
                    $compCyclesMembers = (Get-CMCollectionMember -CollectionName $deviceColl).Name | Sort-Object

                    Append-ColoredLine -box $outputBox2 -color DarkBlue -text "$deviceColl`r`n" -bold $true -fontSize 15 -textAlign Center
                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nCompliance Cycles:`r`n" -bold $true -fontSize 13 -textAlign Left
                    
                    if (![string]::IsNullOrWhiteSpace($compCyclesMembers)) {
                        # Create 'Device Collection' in 'Compliance Cycles' folder if not present
                        if ((Test-Path "$path\Compliance Cycles") -eq $false)
                        { New-Item -ItemType Directory -Name "Compliance Cycles" -Path $path | Out-Null }
                        
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nTotal Count = $($compCyclesMembers.Count)`r" -bold $false -fontSize 11 -textAlign Left
                        Append-ColoredLine -box $outputBox2 -color Black -text "`nRunning cycles now`r`n" -bold $false -fontSize 11 -textAlign Left

                        $cyclesReportSuc, $cyclesReportFai = Run-Cycles -serversList $compCyclesMembers -errorDisplay $false

                        if (![string]::IsNullOrWhiteSpace($cyclesReportSuc)) {
                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`nSuccessfully executed on the following servers: (Count = $($cyclesReportSuc.Count)`)`r`n" -bold $true -fontSize 11 -textAlign Left
                            Add-OutputBoxLine -Message ($cyclesReportSuc | Out-String) -scrollCheck $false -fontSize2 11 -lucidaBool $true
                        }

                        if (![string]::IsNullOrWhiteSpace($cyclesReportFai)) {
                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`nFailed on the following servers: (Count = $($cyclesReportFai.Count)`)`r`n" -bold $true -fontSize 11 -textAlign Left
                            Add-OutputBoxLine -Message ($cyclesReportFai | Out-String) -scrollCheck $false -fontSize2 11 -lucidaBool $true
                        }
                                              
                        "==========Compliance Cycles==========", "", "User: $scriptUserName", "Date: $date", "", "Device Collection: $deviceColl" | Out-File "$path\Compliance Cycles\Device Collection\compliance_cycles - $dateTime.txt" -Append

                        if ($null -ne $cyclesReportSuc) { "", "Success:", $cyclesReportSuc | Out-File "$path\Compliance Cycles\Device Collection\compliance_cycles - $dateTime.txt" -Append }
                        if ($null -ne $cyclesReportFai) { "", "Failed:", $cyclesReportFai | Out-File "$path\Compliance Cycles\Device Collection\compliance_cycles - $dateTime.txt" -Append }
                        Set-ItemProperty -Path "$path\Compliance Cycles\Device Collection\compliance_cycles - $dateTime.txt" -Name IsReadOnly -Value $true

                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nLogs saved at the following path:" -bold $true -fontSize 11 -textAlign Left
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nCompliance Cycles\Device Collection\compliance_cycles - $dateTime.txt" -bold $false -fontSize 11 -textAlign Left
                    }
                    else { Append-ColoredLine -box $outputBox2 -color Red -text "`r`nEmpty Collection" -bold $true -fontSize 11 -textAlign Left }
                }

                # Compliance Check
                if ($radioCompCheck.Checked) {
                    $yesButtonEmptyPatch.Visible = $false
                    $yesButtonFullPatch.Visible = $false
                    $yesButtonReboot.Visible = $false
                    $noButtonEmptyPatch.Visible = $false
                    $noButtonFullPatch.Visible = $false
                    $noButtonReboot.Visible = $false

                    $compliant = @()
                    Append-ColoredLine -box $outputBox2 -color DarkBlue -text "$deviceColl`r`n" -bold $true -fontSize 15 -textAlign Center
                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nCompliance Check:`r`n" -bold $true -fontSize 13 -textAlign Left
                    $compCheckMembers = (Get-CMCollectionMember -CollectionName $deviceColl).Name | Sort-Object

                    if (![string]::IsNullOrWhiteSpace($compCheckMembers)) {
                        <#
                        if($checkQuickCompCheck.Checked -ne $true)
                        {                       
                            # CYCLES
                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`nTotal Count = $($compCheckMembers.Count)`r" -bold $true -fontSize 11 -textAlign Left
                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`nRunning cycles now | " -bold $false -fontSize 11 -textAlign Left
                            #Add-OutputBoxLine -Message "`r`nTotal Count = $($compCheckMembers.Count)`r" -scrollCheck $true -fontSize2 11 -lucidaBool $false
                            #Add-OutputBoxLine -Message "`r`nRunning cycles now | " -scrollCheck $true -fontSize2 11 -lucidaBool $false

                            #$cyclesRepSuc, $cyclesRepFai = Run-Cycles -serversList $compCheckMembers -errorDisplay $false
                            #Start-Countdown -Seconds $timeS -Message "Running compliance cycles"
                            Append-ColoredLine -box $outputBox2 -color Green -text "done`r" -bold $false -fontSize 11 -textAlign Left
                        }#>

                        
                        # Create 'Device Collection' in 'Compliance Check' folder if not present
                        if ((Test-Path "$path\Compliance Check\Device Collection") -eq $false)
                        { New-Item -ItemType Directory -Name "Device Collection" -Path "$path\Compliance Check" | Out-Null }

                        # DEPLOYMENT SUMMARIZATION
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nInitiated deployment summarization, waiting 60 seconds`r" -bold $false -fontSize 11 -textAlign Left
                        Invoke-CMDeploymentSummarization -CollectionName $deviceColl
                        Start-Countdown -Seconds $timeD -Message "Waiting for deployment summarization to run"

                        <#
                        #$outputBox2.Controls.Add($timerLabel)
                        $outputBox2.Text = ("`r`n`nSeconds Remaining: ")
                        
                        while ($timeD -ge 1)
                        {
                            #Append-ColoredLine -box $outputBox2 -color Black -text "`r`n`nSeconds Remaining: $($timeD)" -bold $false -fontSize 11 -textAlign Left
                            #Add-OutputBoxLine -Message "`r`n`nSeconds Remaining: $($timeD)" -scrollCheck $true -fontSize2 11 -lucidaBool $false
                            
                            $outputBox2.Text = $outputBox2.Text.Replace($timeD, $timeD-1)

                            #$timerLabel.Text = "Seconds Remaining: $($timeD)"
                            Start-Sleep 1
                            $timeD -= 1
                        }
                        while ($delay -ge 1)
                        {
                            $Counter_box.Refresh()
                            $Counter_box.Text = $Counter_box.Text.Replace("$delay", $delay-1)
                            #$Counter_box.appendText($delay)
                            start-sleep 1
                            $delay -= 1
                        }
                        #>
                        #Append-ColoredLine -box $outputBox2 -color Black -text "`r`nDone`r" -bold $false -fontSize 11 -textAlign Left
                        
                        # COMPLIANCE
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nChecking Compliance Status`r`n" -bold $false -fontSize 11 -textAlign Left
                        #Add-OutputBoxLine -Message "`r`nChecking Compliance Status`r`n" -scrollCheck $false -fontSize2 11 -lucidaBool $false
                        
                        foreach ($compCheckMem in $compCheckMembers)
                        { $compliant += Get-SCCMSoftwareUpdateStatus -CollectionName $deviceColl -Server $compCheckMem -checkValidDeployment $true }
                        
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nCompliance Status:`r`n`n" -bold $true -fontSize 11 -textAlign Left

                        $compliantSuccess = ($compliant | Where-Object { $null -ne $_.CollectionName } | Sort-Object -Property Status | Out-String).Trim()
                        #Append-ColoredLine -box $outputBox2 -color Black -text "`r`n$compliantSuccess" -bold $false -fontSize 10
                        Add-OutputBoxLine -Message ($compliantSuccess | Out-String) -scrollCheck $false -fontSize2 13 -lucidaBool $true

                        $compliantFailed = $compliant | Where-Object { $null -eq $_.CollectionName }

                        "==========Compliance Check==========", "", "User: $scriptUserName", "Date: $date", "", "Device Collection: $deviceColl", "" | Out-File "$path\Compliance Check\Device Collection\compliance_check - $dateTime.txt" -Append

                        if ($null -ne $compliantSuccess) { $compliantSuccess | Sort-Object -Property Status | Out-File "$path\Compliance Check\Device Collection\compliance_check - $dateTime.txt" -Append }
                        #if(![string]::IsNullOrWhiteSpace($compliantFailed)) 
                        if ($null -ne $compliantFailed) {
                            "", "Failed to Check:" | Out-File "$path\Compliance Check\Device Collection\compliance_check - $dateTime.txt" -Append
                            $compliantFailed | Out-File "$path\Compliance Check\Device Collection\compliance_check - $dateTime.txt" -Append
                        }

                        Set-ItemProperty -Path "$path\Compliance Check\Device Collection\compliance_check - $dateTime.txt" -Name IsReadOnly -Value $true

                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nLogs saved at the following path:`r" -bold $true -fontSize 11 -textAlign Left
                        Append-ColoredLine -box $outputBox2 -color Black -text "Compliance Check\Device Collection\compliance_check - $dateTime.txt`r" -bold $false -fontSize 11 -textAlign Left
                        #Append-ColoredLine -box $outputBox2 -color Black -text "`r`nLogs saved at the following path:" -bold $true -fontSize 11 -textAlign Left
                        #Add-OutputBoxLine -Message "`r`nCompliance Check\compliance_check - $dateTime.txt" -scrollCheck $true -fontSize2 11 -lucidaBool $false
                        #>
                    }
                    else { Append-ColoredLine -box $outputBox2 -color Red -text "`rEmpty Collection" -bold $true -fontSize 11 -textAlign Left }
                }
                
                # Patching
                if ($radioPatching.Checked) {
                    $yesButtonEmptyPatch.Visible = $false
                    $yesButtonFullPatch.Visible = $false
                    $yesButtonReboot.Visible = $false
                    $noButtonEmptyPatch.Visible = $false
                    $noButtonFullPatch.Visible = $false
                    $noButtonReboot.Visible = $false

                    $global:finalSug = $textSugName.Text.Trim().ToUpper()
                    $sugCheck = Get-CMSoftwareUpdateGroup -Name $finalSug

                    # Create 'Patch Push' folder if not present
                    if ((Test-Path "$path\Patch Push") -eq $false)
                    { New-Item -ItemType Directory -Name "Patch Push" -Path $path | Out-Null }

                    Append-ColoredLine -box $outputBox2 -color DarkBlue -text "$deviceColl`r`n" -bold $true -fontSize 15 -textAlign Center
                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nApply Patches:`r`n" -bold $true -fontSize 13 -textAlign Left
                    
                    if ($null -ne $sugCheck) {
                        #$outputBox2.Controls.Add($patchingYesButton)
                        #$outputBox2.Controls.Add($patchingNoButton)

                        $global:patchMembers = (Get-CMCollectionMember -CollectionName $deviceColl).Name | Sort-Object
                    
                        # If Empty Device Collection
                        if ([string]::IsNullOrWhiteSpace($patchMembers)) {
                            Append-ColoredLine -box $outputBox2 -color Red -text "`r`nEmpty Collection`r" -bold $true -fontSize 11 -textAlign Left
                        
                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`n`nPatches from $finalSug will be pushed on this empty device collection. Do you wish to continue?`r" -bold $true -fontSize 11 -textAlign Left
                        
                            $yesButtonEmptyPatch.Visible = $true
                            $noButtonEmptyPatch.Visible = $true

                            $yesButtonEmptyPatch.Add_Click(
                                {
                                    $yesButtonEmptyPatch.Visible = $false
                                    $yesButtonFullPatch.Visible = $false
                                    $yesButtonReboot.Visible = $false
                                    $noButtonEmptyPatch.Visible = $false
                                    $noButtonFullPatch.Visible = $false
                                    $noButtonReboot.Visible = $false

                                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nApplying $finalSug to $deviceColl" -bold $false -fontSize 11 -textAlign Left
                            
                                    New-CMSoftwareUpdateDeployment -DeploymentName "Microsoft Software Updates - $date" -Description "Maker: $($dropMaker.Text) Checker: $($dropChecker.Text) Date: $date" -SoftwareUpdateGroupName $finalSug -CollectionName $($dropDCList.Text) -DeploymentType Required -VerbosityLevel OnlySuccessAndErrorMessages -TimeBasedOn LocalTime -AvailableDateTime $date -DeadlineDateTime $date -UserNotification DisplayAll -SoftwareInstallation $true -RestartServer $true -RequirePostRebootFullScan $true -PersistOnWriteFilterDevice $true -ProtectedType RemoteDistributionPoint -UnprotectedType UnprotectedDistributionPoint -AcceptEula:$true | Out-Null -ErrorAction Stop
                            
                                    Append-ColoredLine -box $outputBox2 -color Green -text "`rSUG was applied successfully`r" -bold $false -fontSize 11 -textAlign Left
                            
                                    "==========Patch Push==========", "", "Date: $date", "Maker: $($dropMaker.Text)", "Checker: $($dropChecker.Text)", `
                                        "", "Device Collection: $($dropDCList.Text)", "SUG: $finalSug", "", "Members: Empty Device Collection" | Out-File "$path\Patch Push\patch_push - $dateTime.txt" -Append
                            
                                    Set-ItemProperty -Path "$path\Patch Push\patch_push - $dateTime.txt" -Name IsReadOnly -Value $true

                                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nLogs saved at the following path:`r" -bold $true -fontSize 11 -textAlign Left
                                    Append-ColoredLine -box $outputBox2 -color Black -text "Patch Push\patch_push - $dateTime.txt`r" -bold $false -fontSize 11 -textAlign Left

                                })
                            #> # Yes Button

                            #$yesButton.Add_Click({Confirm-Patching -empty $true})
                        
                            #Append-ColoredLine -box $outputBox2 -color Red -text "`r`n`nCancelled" -bold $true -fontSize 10
                            #}

                            $noButtonEmptyPatch.Add_Click({
                                    #$counter ++
                                    Append-ColoredLine -box $outputBox2 -color Red -text "`r`nCancelled" -bold $true -fontSize 11 -textAlign Left
                                    #$outputBox2.Clear()
                                    $yesButtonEmptyPatch.Visible = $false
                                    $yesButtonFullPatch.Visible = $false
                                    $yesButtonReboot.Visible = $false
                                    $noButtonEmptyPatch.Visible = $false
                                    $noButtonFullPatch.Visible = $false
                                    $noButtonReboot.Visible = $false
                                    #$yesButton.Enabled = $false
                                    #$noButton.Enabled = $false
                            
                                    $form.Refresh()

                                })
                        }

                        # If Not Empty Device Collection
                        if (![string]::IsNullOrWhiteSpace($patchMembers)) {
                            #Add-OutputBoxLine -Message "`r`nDevice Collection has following servers: (Count - $($patchMembers.Count)`)" -scrollCheck $true -fontSize2 10
                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`nDevice Collection has following servers: (Count = $($patchMembers.Count)`)`r" -bold $true -fontSize 11 -textAlign Left
                            Add-OutputBoxLine -Message ($patchMembers | Out-String) -scrollCheck $false  -fontSize2 11 -lucidaBool $true
                        
                            #$confirmPatch = [System.Windows.MessageBox]::Show('Patches from ' + $finalSug + ' will be pushed on the displayed servers. Do you wish to continue?', 'Confirmation', 'YesNo', 'Warning')
                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`nPatches from $finalSug will be pushed on the displayed servers. Do you wish to continue?`r" -bold $true -fontSize 11 -textAlign Left
                        
                            $yesButtonFullPatch.Visible = $true
                            $noButtonFullPatch.Visible = $true
                        
                            $yesButtonFullPatch.Add_Click({

                                    # APPLY PATCHES
                                    $yesButtonEmptyPatch.Visible = $false
                                    $yesButtonFullPatch.Visible = $false
                                    $yesButtonReboot.Visible = $false
                                    $noButtonEmptyPatch.Visible = $false
                                    $noButtonFullPatch.Visible = $false
                                    $noButtonReboot.Visible = $false
                                    #>

                                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nApplying $finalSug to $deviceColl" -bold $false -fontSize 11 -textAlign Left
                            
                                    New-CMSoftwareUpdateDeployment -DeploymentName "Microsoft Software Updates - $date" -Description "Maker: $($dropMaker.Text) Checker: $($dropChecker.Text) Date: $date" -SoftwareUpdateGroupName $finalSug -CollectionName $($dropDCList.Text) -DeploymentType Required -VerbosityLevel OnlySuccessAndErrorMessages -TimeBasedOn LocalTime -AvailableDateTime $date -DeadlineDateTime $date -UserNotification DisplayAll -SoftwareInstallation $true -RestartServer $true -RequirePostRebootFullScan $true -PersistOnWriteFilterDevice $true -ProtectedType RemoteDistributionPoint -UnprotectedType UnprotectedDistributionPoint -AcceptEula:$true | Out-Null -ErrorAction Stop
                            
                                    Append-ColoredLine -box $outputBox2 -color Green -text "`rSUG was applied successfully`r" -bold $false -fontSize 11 -textAlign Left

                                    "==========Patch Push==========", "", "Date: $date", "Maker: $($dropMaker.Text)", "Checker: $($dropChecker.Text)", `
                                        "", "Device Collection: $($dropDCList.Text)", "SUG: $finalSug", "", "Members: ", $patchMembers | Out-File "$path\Patch Push\patch_push - $dateTime.txt" -Append
                            
                                    Set-ItemProperty -Path "$path\Patch Push\patch_push - $dateTime.txt" -Name IsReadOnly -Value $true

                                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nLogs saved at the following path:`r`n" -bold $true -fontSize 11 -textAlign Left
                                    Append-ColoredLine -box $outputBox2 -color Black -text "Patch Push\patch_push - $dateTime.txt`r" -bold $false -fontSize 11 -textAlign Left
                                })#>
                        
                            $noButtonFullPatch.Add_Click({
                                    Append-ColoredLine -box $outputBox2 -color Red -text "`r`nCancelled" -bold $true -fontSize 11 -textAlign Left
                                    $yesButtonEmptyPatch.Visible = $false
                                    $yesButtonFullPatch.Visible = $false
                                    $yesButtonReboot.Visible = $false
                                    $noButtonEmptyPatch.Visible = $false
                                    $noButtonFullPatch.Visible = $false
                                    $noButtonReboot.Visible = $false
                                    #>
                                })
                        }
                    
                    } # SUG Check
                    
                    else { Append-ColoredLine -box $outputBox2 -color Red -text "`r`nNot a valid SUG" -bold $true -fontSize 11 -textAlign Left }
                }

                # Reboot
                if (($radioReboot.Checked) -and ($radioReboot.Enabled)) {
                    $yesButtonEmptyPatch.Visible = $false
                    $yesButtonFullPatch.Visible = $false
                    $yesButtonReboot.Visible = $false
                    $noButtonEmptyPatch.Visible = $false
                    $noButtonFullPatch.Visible = $false
                    $noButtonReboot.Visible = $false
                
                    Append-ColoredLine -box $outputBox2 -color DarkBlue -text "$deviceColl`r`n" -bold $true -fontSize 15 -textAlign Center
                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nReboot Servers:`r`n" -bold $true -fontSize 13 -textAlign Left
                    
                    # Defining arrays
                    $global:connFailed = @()     # Connection Failed
                    $global:rebootDefault = @()  # Servers that require 1st reboot without running cycles
                    $rebootStatus = @()
                    $rebootStatusFailed = @()
                    $rebootCheck2 = @()

                    $firstReboot = @()    # Servers that require 1st reboot after running cycles
                    $secondReboot = @()   # Servers that require 2nd reboot
                    $thirdReboot = @()    # Servers that require 3rd reboot

                    $zeroCompliant = @()  # Compliant without any reboot
                    $firstCompliant = @() # Compliant after 1st reboot
                    $secondCompliant = @()# Compliant after 2nd reboot
                    $compliant3 = @()
                   
                    # Get Device Collection Members
                    $global:rebootMembers = (Get-CMCollectionMember -CollectionName $deviceColl).Name | Sort-Object
                    #>
                    if (![string]::IsNullOrEmpty($rebootMembers)) {
                        # Create 'Reboot Status' folder if not present
                        if ((Test-Path "$path\Reboot & Compliance") -eq $false)
                        { New-Item -ItemType Directory -Name "Reboot & Compliance" -Path $path | Out-Null }

                        # Create 'Connection Failed' folder if not present
                        if ((Test-Path "$path\Reboot & Compliance\Connection Failed") -eq $false)
                        { New-Item -ItemType Directory -Name "Connection Failed" -Path "$path\Reboot & Compliance" | Out-Null }
                        
                        # Check if reboot is required on all servers using REGISTRY
                        #Add-OutputBoxLine -Message "Checking reboot status of the servers" -scrollCheck $true
                        #Add-OutputBoxLine -Message "`r`nCount: $($rebootMembers.Count)" -scrollCheck $true

                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nChecking reboot status of the servers: (Count = $($rebootMembers.Count)`)`r" -bold $true -fontSize 11 -textAlign Left

                        foreach ($mem1 in $rebootMembers) {
                            try {
                                $rebootCheck = Invoke-Command -ComputerName $mem1 -ScriptBlock { Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction Ignore } -ErrorAction Stop
                                #$rebootCheck = Invoke-Command -ComputerName $mem1 -ScriptBlock{hostname} -ErrorAction Stop
                                if ($null -ne $rebootCheck) { $global:rebootDefault += $mem1 }
                            }
                            #catch [System.Management.Automation.Remoting.PSRemotingTransportException]
                            catch {
                                $connFailed += $mem1
                            }
                        }

                        # Report unreachable servers
                        if (![string]::IsNullOrEmpty($connFailed)) {

                            #$countConnFailed = $connFailed.Count

                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`n+++++++++++++++++++++++++++++++`r" -bold $true -fontSize 11 -textAlign Left
                            
                            #Add-OutputBoxLine -Message "`rConnection failed on following servers:`r`n`n" -scrollCheck $true
                            #Append-ColoredLine -box $outputBox2 -color Black -text "Connection failed on following servers:`r" -bold $true -fontSize 10
                            #Add-OutputBoxLine -Message "Count: $($connFailed.Count)`r`n`n" -scrollCheck $true -fontSize2 10
                            
                            Append-ColoredLine -box $outputBox2 -color Black -text "Connection failed on following servers: (Count = $($connFailed.Count)`)`r`n`n" -bold $true -fontSize 11 -textAlign Left
                            Add-OutputBoxLine -Message (($connFailed -join ", " | Out-String).Trim()) -scrollCheck $false -fontSize2 11 -lucidaBool $true
                            
                            #Append-ColoredLine -box $outputBox2 -color Red -text ($connFailed | Out-String) -bold $false -fontSize 10 -textAlign Left

                            "==========Reboot & Compliance==========", "", "Date: $date", "Maker: $($dropMaker.Text)", "Checker: $($dropChecker.Text)", `
                                "Device Collection: $($dropDCList.Text)", "", "Connection failed on following servers:", $connFailed | Out-File "$path\Reboot & Compliance\Connection Failed\reboot_unreachable - $dateTime.txt" -Append
                            
                            Set-ItemProperty -Path "$path\Reboot & Compliance\Connection Failed\reboot_unreachable - $dateTime.txt" -Name IsReadOnly -Value $true

                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`nLogs saved at the following path:`r" -bold $true -fontSize 11 -textAlign Left
                            Append-ColoredLine -box $outputBox2 -color Black -text "Reboot & Compliance\Connection Failed\reboot_unreachable - $dateTime.txt`r" -bold $false -fontSize 11 -textAlign Left
                            #Add-OutputBoxLine -Message "+++++++++++++++++++++++++++++++" -scrollCheck $true
                            Append-ColoredLine -box $outputBox2 -color Black -text "+++++++++++++++++++++++++++++++`r" -bold $false -fontSize 12 -textAlign Left
                            
                            #$compared = Compare-Object $rebootDefault $connFailed -IncludeEqual
                            #if(($compared.Sideindicator -contains "<=") -and ($compared.Sideindicator -contains "=>"))
                            #{Append-ColoredLine -box $outputBox2 -color Red -text "`r`nCould not check on any server`rMake sure to use CyberArk for production servers" -bold $true -fontSize 10
                            #}
                        }

                        # $rebootDefault = reboot required servers
                        # $connFailed = reboot check failed servers
                        
                        # No server require a reboot, connected on all
                        if (([string]::IsNullOrEmpty($rebootDefault)) -and ([string]::IsNullOrEmpty($connFailed)))
                        { Append-ColoredLine -box $outputBox2 -color Green -text "`r`nNone of the servers require a reboot" -bold $true -fontSize 11 -textAlign Left }

                        # No connection on all servers
                        if (([string]::IsNullOrEmpty($rebootDefault)) -and ($($rebootMembers.Count) -eq $($connFailed.Count)))
                        { Append-ColoredLine -box $outputBox2 -color Red -text "`r`nCould not check on any server`rMake sure to use CyberArk for production servers" -bold $true -fontSize 11 -textAlign Left }

                        # No reboot required on connected servers, includes connection failed servers
                        if (([string]::IsNullOrEmpty($rebootDefault)) -and (![string]::IsNullOrEmpty($connFailed) -and ($($rebootMembers.Count) -ne $($connFailed.Count))))
                        { Append-ColoredLine -box $outputBox2 -color Green -text "`r`nNone of the servers on which connection was established require a reboot" -bold $true -fontSize 11 -textAlign Left }

                        # If servers require reboot
                        if (![string]::IsNullOrEmpty($rebootDefault)) {

                            #Add-OutputBoxLine -Message "`r`nFollowing servers are pending for reboot:`r`n`n" -scrollCheck $true
                            #Append-ColoredLine -box $outputBox2 -color Black -text "`r`nChecking reboot status of the servers: (Count = $($rebootMembers.Count)`)`r" -bold $true -fontSize 11 -textAlign Left

                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`nFollowing servers are pending for reboot: (Count = $($rebootDefault.Count)`)`r" -bold $true -fontSize 11 -textAlign Left
                            Add-OutputBoxLine -Message ($rebootDefault | Out-String) -scrollCheck $false -fontSize2 12 -lucidaBool $true #hahacheck
                            #Add-OutputBoxLine -Message ($members | Out-String) -scrollCheck $false -fontSize2 12 -lucidaBool $true
                        
                            #$confirmReboot = [System.Windows.MessageBox]::Show('Displayed servers will be rebooted. Do you wish to continue?', 'Confirmation', 'YesNo', 'Warning')
                            Append-ColoredLine -box $outputBox2 -color Black -text "`r`nDisplayed servers will be rebooted. Do you wish to continue?`r" -bold $true -fontSize 11 -textAlign Left
                        
                            $yesButtonReboot.Visible = $true
                            $noButtonReboot.Visible = $true    
                                                                    
                            #if($confirmReboot -eq "Yes")
                            $yesButtonReboot.Add_Click({ 
                                    $yesButtonEmptyPatch.Visible = $false
                                    $yesButtonFullPatch.Visible = $false
                                    $yesButtonReboot.Visible = $false
                                    $noButtonEmptyPatch.Visible = $false
                                    $noButtonFullPatch.Visible = $false
                                    $noButtonReboot.Visible = $false
                                    #================== First: Reboot ===================
                                    Add-OutputBoxLine -Message "`r`nInitiated first reboot. Waiting for 10 minutes...`r`n" -scrollCheck $true -fontSize2 12 -textAlign Left -lucidaBool $false
                                    #Append-ColoredLine -box $outputBox2 -color Black -text "`r`nInitiated first reboot`r`n" -bold $false -fontSize 11 -textAlign Left
                            
                                    #<#
                                    foreach ($reb1 in $rebootDefault) {
                                        try { Invoke-Command -ComputerName $reb1 -ScriptBlock { Restart-Computer -Force } -ErrorAction Stop } 
                                        catch {
                                            Append-ColoredLine -box $outputBox2 -color Red -text "$reb1 - Could not initiate first reboot`r" -bold $false -fontSize 11 -textAlign Left
                                        }
                                    }
                            
                                    Start-Countdown -Seconds $countdownReboot -Message "Rebooting"
                                    #>
                                    Add-OutputBoxLine -Message "`rDone" -scrollCheck $true -fontSize2 12 -textAlign Left -lucidaBool $false
                                    #Append-ColoredLine -box $outputBox2 -color Black -text "`rDone" -bold $false -fontSize 11 -textAlign Left

                                    #================== Second: Reboot ===================

                                    Add-OutputBoxLine -Message "`r`n`nInitiated second reboot. Waiting for 10 minutes...`r`n" -scrollCheck $true -fontSize2 12 -textAlign Left -lucidaBool $false
                                    #Append-ColoredLine -box $outputBox2 -color Black -text "`r`n`nInitiated second reboot`r`n" -bold $false -fontSize 11 -textAlign Left
                                    #<#
                                    foreach ($reb2 in $rebootDefault) {
                                        try { Invoke-Command -ComputerName $reb2 -ScriptBlock { Restart-Computer -Force } -ErrorAction Stop }
                                        catch {
                                            #Write-Host "$reb2 | Could not initiate second reboot" -ForegroundColor Red
                                            Append-ColoredLine -box $outputBox2 -color Red -text "$reb2 - Could not initiate second reboot`r" -bold $false -fontSize 11 -textAlign Left
                                        }
                                    }
                            
                                    Start-Countdown -Seconds $countdownReboot -Message "Initiated second reboot"
                                    #>

                                    Add-OutputBoxLine -Message "`rDone" -scrollCheck $true -fontSize2 12 -textAlign Left -lucidaBool $false
                                    #Append-ColoredLine -box $outputBox2 -color Black -text "`rDone" -bold $false -fontSize 11 -textAlign Left

                                    #================== Post Reboot ===================

                                    #Add-OutputBoxLine -Message "`r`n`nReboots completed" -scrollCheck $true
                                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`n`n++++++++++++++++Reboots completed++++++++++++++++`r" -bold $true -fontSize 11 -textAlign Left

                                    #Add-OutputBoxLine -Message "`r`nChecking compliance`r" -scrollCheck $true -fontSize2 12 -textAlign Left -lucidaBool $false

                                    Add-OutputBoxLine -Message "`r`nInitiated Cycles. Waiting for 2 minutes...`r`n" -scrollCheck $true -fontSize2 12 -textAlign Left -lucidaBool $false

                                    $cyclesRebSuc, $cyclesRebFai = Run-Cycles -serversList $rebootDefault
                                    Start-Countdown -Seconds $timeS -Message "Waiting for cycles to run"

                                    Add-OutputBoxLine -Message "`r`nInitiated deployment summarization. Waiting for 1 minute...`r`n" -scrollCheck $true -fontSize2 12 -textAlign Left -lucidaBool $false

                                    Invoke-CMDeploymentSummarization -CollectionName $deviceColl
                                    Start-Countdown -Seconds $timeD -Message "Waiting for deployment summarization to run"

                                    # Third : Compliance check
                                    foreach ($mem2 in $rebootDefault)
                                    { $rebootStatus += Get-SCCMSoftwareUpdateStatus -CollectionName $deviceColl -Server $mem2 -checkValidDeployment $true | Sort-Object -Property Status }

                                    <#
                            foreach($mem4 in $rebootStatus)
                            {
                                try{$rebootCheck2 = Invoke-Command -ComputerName $mem4 -ScriptBlock{Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction Ignore} -ErrorAction Stop
                                    if($rebootCheck2 -ne $null){$rebootDefault2 += $mem4}
                                }catch{}
                            }

                            if(($rebootStatus -ne $null) -or ($rebootCheck2 -ne $null))
                            {
                                $rebootFinal = (Compare-Object $rebootStatus $rebootCheck2 -IncludeEqual | Where-Object {$_.SideIndicator -eq "=="}).InputObject
                            }
                            #>

                                    Add-OutputBoxLine -Message "`r`nChecking compliance`r`n" -scrollCheck $true -fontSize2 12 -textAlign Left -lucidaBool $false

                                    $compliantReb = $rebootStatus | Where-Object { $_.Status -eq "Compliant" } | Sort-Object -Property DeviceName
                                    Add-OutputBoxLine -Message ($compliantReb | Out-String) -scrollCheck $true -fontSize2 11 -lucidaBool $true

                                    $thirdReboot = $rebootStatus | Where-Object { ($_.Status -ne "Compliant") -and ($null -ne $_.CollectionName) } | Sort-Object -Property DeviceName
                                    #if([string]::IsNullOrWhiteSpace($thirdReboot))
                                    if ($null -ne $thirdReboot) {
                                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`n`nFollowing servers require 3rd reboot:`r`n`n" -bold $true -fontSize 11 -textAlign Left 
                                        Add-OutputBoxLine -Message ($thirdReboot | Out-String) -scrollCheck $true -fontSize2 11 -lucidaBool $true
                                    }
                                    <#
                            $rebootStatusFailed = $rebootStatus | Where-Object {$_.CollectionName -eq $null} | Sort-Object  
                            if($null -ne $rebootStatusFailed)
                            {
                                Append-ColoredLine -box $outputBox2 -color Black -text "`r`n`nFailed to check compliance on following servers:`r`n`n" -bold $true -fontSize 11 -textAlign Left
                                Add-OutputBoxLine -Message ($rebootStatusFailed | Out-String) -scrollCheck $true -fontSize2 11 -lucidaBool $true
                            }
                            #>
                                    "==========Reboot & Compliance==========", "", "Date: $date", "Maker: $($dropMaker.Text)", "Checker: $($dropChecker.Text)", `
                                        "Device Collection: $deviceColl" | Out-File "$path\Reboot & Compliance\reboot_compliance - $dateTime.txt" -Append
                            
                                    # Out | Complaint servers
                                    if ($null -ne $compliantReb) { $compliantReb | Sort-Object -Property Status | Out-File "$path\Reboot & Compliance\reboot_compliance - $dateTime.txt" -Append }
                            
                                    # Out | Third Reboot servers
                                    if ($null -ne $thirdReboot) {
                                        "", "Third Reboot:" | Out-File "$path\Reboot & Compliance\reboot_compliance - $dateTime.txt" -Append
                                        $thirdReboot | Out-File "$path\Reboot & Compliance\reboot_compliance - $dateTime.txt" -Append
                                    }

                                    <#
                            # Out | Failed to check compliance servers
                            if($null -ne $rebootStatusFailed) {
                                "","Failed to Check Compliance:" | Out-File "$path\Reboot & Compliance\reboot_compliance - $dateTime.txt" -Append
                                $rebootStatusFailed | Out-File "$path\Reboot & Compliance\reboot_compliance - $dateTime.txt" -Append
                            }
                            #>

                                    # Out | Reboot not required servers
                                    $rebootNotRequired = Compare-Object $rebootMembers $rebootDefault | Where-Object { $_.SideIndicator -eq "<=" } | Select-Object -ExpandProperty InputObject
                                    if ($null -ne $rebootNotRequired) { "", "No reboots required:", $rebootNotRequired | Out-File "$path\Reboot & Compliance\reboot_compliance - $dateTime.txt" -Append }

                                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`n`nRunning compliance cycles on servers that require 3rd reboot`r" -bold $true -fontSize 11 -textAlign Left
                                    $check1, $check2 = Run-Cycles -serversList $thirdReboot
                                    Add-OutputBoxLine -Message "`r`nDone" -scrollCheck $true -fontSize2 12 -textAlign Left -lucidaBool $false

                                    Set-ItemProperty -Path "$path\Reboot & Compliance\reboot_compliance - $dateTime.txt" -Name IsReadOnly -Value $true

                                    #Write-Host "`nReport saved in C:\Script\Akashdeep\Patching\logs\" -NoNewline
                                    #Write-Host "Reboot & Compliance\reboot_compliance - $dateTime.txt" -ForegroundColor Yellow
                                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`n`nLogs saved at the following path:`r" -bold $true -fontSize 11 -textAlign Left
                                    Append-ColoredLine -box $outputBox2 -color Black -text "Reboot & Compliance\reboot_compliance - $dateTime.txt`r" -bold $false -fontSize 11 -textAlign Left

                                })

                            #else{#Write-Host "`nCancelled by user" -ForegroundColor Gray
                            #    Append-ColoredLine -box $outputBox2 -color Red -text "Cancelled" -bold $true -fontSize 10
                            #}                        
                        
                            $noButtonReboot.Add_Click({
                                    Append-ColoredLine -box $outputBox2 -color Red -text "`r`nCancelled" -bold $true -fontSize 11 -textAlign Left
                                    $yesButtonEmptyPatch.Visible = $false
                                    $yesButtonFullPatch.Visible = $false
                                    $yesButtonReboot.Visible = $false
                                    $noButtonEmptyPatch.Visible = $false
                                    $noButtonFullPatch.Visible = $false
                                    $noButtonReboot.Visible = $false
                                })

                        }

                        #$compared = Compare-Object $rebootDefault $connFailed -IncludeEqual

                        
                        <#
                        # If could not connect on any servers
                        if(($rebootDefault -eq $null) -and (($compared.Sideindicator -notcontains "<=") -and ($compared.Sideindicator -notcontains "=>")))
                        {
                            #Write-Host "None of the servers require a reboot" -ForegroundColor Green
                            #Append-ColoredLine -box $outputBox2 -color Green -text "`r`nNone of the servers require a reboot" -bold $true -fontSize 10
                            Append-ColoredLine -box $outputBox2 -color Red -text "`r`nCould not check on any server`rMake sure to use CyberArk for production servers" -bold $true -fontSize 10
                        } # Without running cycles
                        #>
                    }
                    else {
                        #Write-Host "Empty Device Collection" -ForegroundColor Gray
                        #Add-OutputBoxLine -Message "Empty Collection" -scrollCheck $true
                        Append-ColoredLine -box $outputBox2 -color Red -text "`r`nEmpty Collection" -bold $true -fontSize 11 -textAlign Left
                    }

                } # If Radio button of reboot checked
                
            }
            else {
                #Write-Host "Device collection does not exist" -ForegroundColor Magenta
                #Add-OutputBoxErrorLine -MessageError "Device collection does not exist"
                Append-ColoredLine -box $outputBox2 -color Red -text "`r`nDevice collection does not exist" -bold $true -fontSize 11 -textAlign Left
            }
            #}comment this
        }
    }

    Function appservers-code {
        <#
    $yesButtonEmptyPatch.Visible = $false
    $yesButtonFullPatch.Visible = $false
    $yesButtonReboot.Visible = $false
    $noButtonEmptyPatch.Visible = $false
    $noButtonFullPatch.Visible = $false
    $noButtonReboot.Visible = $false
    #>
        if ([string]::IsNullOrWhiteSpace($textServers.Text))
        { Add-OutputBoxErrorLine -MessageError "Enter some servers" }

        if (![string]::IsNullOrWhiteSpace($textServers.Text)) {
            if (!($radioCompCycServers.Checked) -and !($radioCompCheckServers.Checked))
            { Add-OutputBoxErrorLine -MessageError "Select an action" }

            else {

                #$totalAppServers = ($textServers.Lines).Trim()
                $totalAppServers = $textServers.Lines | Where-Object { !([string]::IsNullOrWhiteSpace($_)) } | Sort-Object
                $textServers.Lines = $totalAppServers

                # Compliance Cycles
                if ($radioCompCycServers.Checked) {
                    #$compCyclesMembers = (Get-CMCollectionMember -CollectionName $deviceColl).Name | Sort-Object
            

                    #Append-ColoredLine -box $outputBox2 -color DarkBlue -text "$deviceColl`r`n" -bold $true -fontSize 15 -textAlign Center
                    Append-ColoredLine -box $outputBox2 -color DarkBlue -text "Compliance Cycles`r`n" -bold $true -fontSize 15 -textAlign Center
                    
                    # Create 'Device Collection' in 'Compliance Cycles' folder if not present
                    if ((Test-Path "$path\Compliance Cycles\Individual Servers") -eq $false)
                    { New-Item -ItemType Directory -Name "Individual Servers" -Path "$path\Compliance Cycles" | Out-Null }
                        
                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nTotal Count = $($totalAppServers.Count)`r" -bold $false -fontSize 11 -textAlign Left
                    Append-ColoredLine -box $outputBox2 -color Black -text "`nRunning cycles now`r`n" -bold $false -fontSize 11 -textAlign Left

                    $cyclesReportSuc, $cyclesReportFai = Run-Cycles -serversList $totalAppServers -errorDisplay $false

                    if (![string]::IsNullOrWhiteSpace($cyclesReportSuc)) {
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nSuccessfully executed on the following servers: (Count = $($cyclesReportSuc.Count)`)`r`n" -bold $true -fontSize 11 -textAlign Left
                        Add-OutputBoxLine -Message ($cyclesReportSuc | Out-String) -scrollCheck $false -fontSize2 11 -lucidaBool $true
                    }

                    if (![string]::IsNullOrWhiteSpace($cyclesReportFai)) {
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nFailed on the following servers: (Count = $($cyclesReportFai.Count)`)`r`n" -bold $true -fontSize 11 -textAlign Left
                        Add-OutputBoxLine -Message ($cyclesReportFai | Out-String) -scrollCheck $false -fontSize2 11 -lucidaBool $true
                    }
                                              
                    "==========Compliance Cycles==========", "", "User: $scriptUserName", "Date: $date" | Out-File "$path\Compliance Cycles\Individual Servers\compliance_cycles - $dateTime.txt" -Append

                    if ($null -ne $cyclesReportSuc) { "", "Success:", $cyclesReportSuc | Out-File "$path\Compliance Cycles\Individual Servers\compliance_cycles - $dateTime.txt" -Append }
                    if ($null -ne $cyclesReportFai) { "", "Failed:", $cyclesReportFai | Out-File "$path\Compliance Cycles\Individual Servers\compliance_cycles - $dateTime.txt" -Append }
                    Set-ItemProperty -Path "$path\Compliance Cycles\Individual Servers\compliance_cycles - $dateTime.txt" -Name IsReadOnly -Value $true

                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nLogs saved at the following path:" -bold $true -fontSize 11 -textAlign Left
                    Append-ColoredLine -box $outputBox2 -color Black -text "`r`nCompliance Cycles\Individual Servers\compliance_cycles - $dateTime.txt" -bold $false -fontSize 11 -textAlign Left
                }
        
                if ($radioCompCheckServers.Checked) {

                    #$global:countdownTime = $null
                    #comehere
        
                    Append-ColoredLine -box $outputBox2 -color DarkBlue -text "Compliance Check`r`n" -bold $true -fontSize 15 -textAlign Center
            
                    # Create 'Device Collection' in 'Compliance Check' folder if not present
                    if ((Test-Path "$path\Compliance Check\Individual Servers") -eq $false)
                    { New-Item -ItemType Directory -Name "Individual Servers" -Path "$path\Compliance Check" | Out-Null }

                    $appCompliant = @()
                    $Finalinput = @()
                    $finalMembers = @()
                    $finalNonMembers = @()

                    $site_server = "TVMSCMP01.belgianrail.be"
                    $site_code = "P01"

                    #$servers=$inputcsv.ServerName
                    #$DCs = Get-content "C:\Script\Akashdeep\Patching\New folder\Individual DC List.txt"
                    #$DCs = $allDeviceCollection
                    if ([string]::IsNullOrWhiteSpace($dropDCCompliance.Text)) {
                        foreach ($global:server in $($textServers.Lines.Trim())) {
                            if (!([string]::IsNullOrWhiteSpace($server))) {
                                $global:server = $server.ToUpper()
                                $collectionname = $null
                                $Collections = (Get-CimInstance -ComputerName $site_server -Namespace root\SMS\site_$site_code `
                                        -Query "SELECT SMS_Collection.* FROM SMS_FullCollectionMembership, SMS_Collection where name = '$server' and SMS_FullCollectionMembership.CollectionID = SMS_Collection.CollectionID").Name | 
                                #Where-Object {($_ -ne "All systems") -and ($_ -ne "All Desktop and Server Clients")}
                                Where-Object { $_ -notmatch "all" }

                                #foreach($dc in $DCs)
                                foreach ($dc in $allDeviceCollection) {
                                    #$dc
                                    $CollectionName = ($Collections | Select-String $dc)
                                    #Append-ColoredLine -box $outputBox2 -color Red -text "`r`nAll Collections Name: $CollectionName`r`n" -bold $false -fontSize 13 -textAlign Left
                                    if ($null -ne $CollectionName) { break }
                                }

                
                                if ($null -eq $collectionname) {
                                    $collectionname = "Not in a patching device collection"
                                    #Append-ColoredLine -box $outputBox2 -color Red -text "`r`n$server is not a part of any Individual DC`r`n" -bold $false -fontSize 13 -textAlign Left
                                    #$server | Out-file "$path\Compliance Check\Individual Servers\Missing from DC\missing_from_dc - $dateTime.txt" -Append
                                }
                                #>
                
                                $data = [ordered]@{
                                    ServerName     = $server
                                    CollectionName = $collectionname
                                }

                                $obj = New-Object -TypeName PSObject -property $data
                                $Finalinput += $obj
                            }
                
                        }
                        $Finalinput = $Finalinput | Sort-Object -property ServerName
                        Add-OutputBoxLine -Message ($Finalinput | Out-String) -scrollCheck $false -fontSize2 11 -lucidaBool $true

                        #Find Unique Device collections given in input and run deployment summarizations on them
                        $UniqueDCs = $Finalinput.CollectionName | Get-unique 
                        #Add-OutputBoxLine -Message ($UniqueDCs | Out-String) -scrollCheck $false -fontSize2 11 -lucidaBool $true
                        #Add-OutputBoxLine -Message ("Running deployment summarization on the above mentioned DCs") -scrollCheck $false -fontSize2 11 -lucidaBool $true
                    }

                    if (![string]::IsNullOrWhiteSpace($dropDCCompliance.Text)) {
                        $global:actualMembers = (Get-CMCollectionMember -CollectionName $($dropDCCompliance.Text)).Name | Sort-Object
                        foreach ($k in $totalAppServers) {
                            #$k
                            if ($k -in $actualMembers) { $finalMembers += $k }
                            else { $finalNonMembers += $k }
                        }

                        if (![string]::IsNullOrWhiteSpace($finalMembers)) {
                            $UniqueDCs = $dropDCCompliance.Text
                            if (![string]::IsNullOrWhiteSpace($finalNonMembers)) {
                                Append-ColoredLine -box $outputBox2 -color Black -text "`r`nFollowing servers are not part of the device collection:`r`n" -bold $false -fontSize 12 -textAlign Left
                                Add-OutputBoxLine -Message ($finalNonMembers | Out-String) -scrollCheck $false -fontSize2 12 -lucidaBool $true
                                #Add-OutputBoxLine -Message ($cyclesReportFai | Out-String) -scrollCheck $false -fontSize2 11 -lucidaBool $true
                            }
                
                        }
                        else { Append-ColoredLine -box $outputBox2 -color Red -text "`r`nNone of the servers exist in this device collection" -bold $true -fontSize 11 -textAlign Left }
                        #else{Add-OutputBoxErrorLine -MessageError "None of the servers exist in this device collection"}
            
                    }

                    if (!([string]::IsNullOrWhiteSpace($UniqueDCs)) -and ($UniqueDCs -ne "Not in a patching device collection")) {
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nInitiating deployment summarization`r" -bold $false -fontSize 12 -textAlign Left
                        #Add-OutputBoxLine -Message ("Running deployment summarization on the above mentioned DCs") -scrollCheck $false -fontSize2 11 -lucidaBool $true
            
                        foreach ($uDC in $UniqueDCs) {
                            if ("Not in a patching device collection" -ne $uDC) {
                                #Add-OutputBoxLine -Message ($uDC | Out-String) -scrollCheck $false -fontSize2 11 -lucidaBool $true
                                Invoke-CMDeploymentSummarization -CollectionName $uDC #removecomment
                            }
                        }
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nWaiting 60 seconds`r" -bold $false -fontSize 12 -textAlign Left
                        Start-Countdown -Seconds $timeD -Message "Waiting for deployment summarization to run" #removecomment
            
                        #endregion: Jobs
            
            
                        # COMPLIANCE
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nChecking Compliance Status`r`n" -bold $false -fontSize 12 -textAlign Left
                        #Add-OutputBoxLine -Message "`r`nChecking Compliance Status`r`n" -scrollCheck $false -fontSize2 11 -lucidaBool $false
                            
                        #foreach($compCheckMem in $compCheckMembers)
                        #{$compliant += Get-SCCMSoftwareUpdateStatus -CollectionName $deviceColl -Server $compCheckMem}
                        if ([string]::IsNullOrWhiteSpace($dropDCCompliance.Text)) {
                            foreach ($i in $finalinput) {
                                $server = $i.ServerName
                                $col = $i.CollectionName
                                $appCompliant += Get-SCCMSoftwareUpdateStatus -CollectionName $col -Server $server -checkValidDeployment $true #removecomment
                            }
                        }
                        if (![string]::IsNullOrWhiteSpace($dropDCCompliance.Text)) {
                            foreach ($global:j in $totalAppServers) {
                                $appCompliant += Get-SCCMSoftwareUpdateStatus -CollectionName $($dropDCCompliance.Text) -Server $j -checkValidDeployment $false #removecomment
                            }
                        }
                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nCompliance Status:`r`n`n" -bold $true -fontSize 11 -textAlign Left

                        $compliantSuccessApp = ($appCompliant | Where-Object { $null -ne $_.CollectionName } | Sort-Object -Property Status | Out-String).Trim()
                        #Append-ColoredLine -box $outputBox2 -color Black -text "`r`n$compliantSuccess" -bold $false -fontSize 10
                        Add-OutputBoxLine -Message ($compliantSuccessApp | Out-String) -scrollCheck $false -fontSize2 13 -lucidaBool $true

                        $compliantFailedApp = $appCompliant | Where-Object { $null -eq $_.CollectionName }

                        "==========Compliance Check==========", "", "User: $scriptUserName", "Date: $date", "", "Device Collection: $deviceColl", "" | Out-File "$path\Compliance Check\Individual Servers\compliance_check - $dateTime.txt" -Append

                        if ($null -ne $compliantSuccessApp) { $compliantSuccessApp | Sort-Object -Property Status | Out-File "$path\Compliance Check\Individual Servers\compliance_check - $dateTime.txt" -Append }
                        #if(![string]::IsNullOrWhiteSpace($compliantFailed)) 
                        if ($null -ne $compliantFailedApp) {
                            "", "Failed to Check:" | Out-File "$path\Compliance Check\Individual Servers\compliance_check - $dateTime.txt" -Append
                            $compliantFailedApp | Out-File "$path\Compliance Check\Individual Servers\compliance_check - $dateTime.txt" -Append
                        }

                        Set-ItemProperty -Path "$path\Compliance Check\Individual Servers\compliance_check - $dateTime.txt" -Name IsReadOnly -Value $true

                        Append-ColoredLine -box $outputBox2 -color Black -text "`r`nLogs saved at the following path:`r" -bold $true -fontSize 11 -textAlign Left
                        Append-ColoredLine -box $outputBox2 -color Black -text "Compliance Check\Individual Servers\compliance_check - $dateTime.txt`r" -bold $false -fontSize 11 -textAlign Left
                        #Append-ColoredLine -box $outputBox2 -color Black -text "`r`nLogs saved at the following path:" -bold $true -fontSize 11 -textAlign Left
                        #Add-OutputBoxLine -Message "`r`nCompliance Check\compliance_check - $dateTime.txt" -scrollCheck $true -fontSize2 11 -lucidaBool $false
                    }
        
                    #>
                }
            }
        }
    }

    $okButtonDC.Add_Click({
            $global:dateTime = Get-Date -Format dd-MM-yyyy_HH-mm-ss
            $global:date = Get-Date
            $outputBox2.Clear()
            $form.Refresh()
            devicecoll-code
        })

    $okButtonServer.Add_Click({
            $global:dateTime = Get-Date -Format dd-MM-yyyy_HH-mm-ss
            $global:date = Get-Date
            $outputBox2.Clear()
            $form.Refresh()
            appservers-code
        })

    $cancelButtonServer.Add_Click({
            # $timer.Stop()
            #$timer.Dispose()
            Write-Host "Cancelled"
        })

    # Set focus to dropdown
    $form.add_shown({ $dropDCList.select() })
    $status = $form.ShowDialog()
    # AD Verification
    else { Write-Host "Not authorized" -ForegroundColor Red }
