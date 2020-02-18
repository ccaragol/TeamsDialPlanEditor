<# 
.Synopsis 
   The purpose of this tool is to give you an easy front end for the management of Tenant Dial
   Plans.  I imagine this may someday be built into the GUI, but until then I hope it helps you out.
 
.DESCRIPTION 
   PowerShell GUI script which allows for GUI management of Tenant Dial Plans
 
.Notes 
     NAME:      Microsoft_Phone_System_Dial_Plan_Editor.ps1
     VERSION:   1.1
     AUTHOR:    C. Anthony Caragol 
     LASTEDIT:  02/18/2020 - Refreshed Code
      
   V 1.1 -	Code Refresh  
	    
.Link 
   Website: http://www.teamsadmin.com
   Twitter: http://www.twitter.com/canthonycaragol
   LinkedIn: http://www.linkedin.com/in/canthonycaragol
 
.EXAMPLE 
   .\Microsoft_Phone_System_Dial_Plan_Editor.ps1

.TODO
    When creating a new dial plan, see if the new name already exists.
    When creating a new dial plan, make sure the input is solid.
    When creating a new dial plan, try a simple error catch and report per PIN setting.
    Copy Normalization Rule to new Dial Plan
    Create Extension Regex
    Change Dial Plan Priority
    Sign the thing with a cert


.APOLOGY
  Please excuse the sloppy coding, I don't use a development environment, IDE or ISE.  I use notepad, 
  not even Notepad++, just notepad.  I am not a developer, just an enthusiast so some code may be redundant or
  inefficient.
#>


$Global:TeamsAdminIcon = [System.Convert]::FromBase64String('
AAABAAEAJiEAAAEACABYCgAAFgAAACgAAAAmAAAAQgAAAAEACAAAAAAAKAUAAAAAAAAAAAAAAAEAAAAB
AAAAAAAAMwAAAGYAAACZAAAAzAAAAP8AAAAAKwAAMysAAGYrAACZKwAAzCsAAP8rAAAAVQAAM1UAAGZV
AACZVQAAzFUAAP9VAAAAgAAAM4AAAGaAAACZgAAAzIAAAP+AAAAAqgAAM6oAAGaqAACZqgAAzKoAAP+q
AAAA1QAAM9UAAGbVAACZ1QAAzNUAAP/VAAAA/wAAM/8AAGb/AACZ/wAAzP8AAP//AAAAADMAMwAzAGYA
MwCZADMAzAAzAP8AMwAAKzMAMyszAGYrMwCZKzMAzCszAP8rMwAAVTMAM1UzAGZVMwCZVTMAzFUzAP9V
MwAAgDMAM4AzAGaAMwCZgDMAzIAzAP+AMwAAqjMAM6ozAGaqMwCZqjMAzKozAP+qMwAA1TMAM9UzAGbV
MwCZ1TMAzNUzAP/VMwAA/zMAM/8zAGb/MwCZ/zMAzP8zAP//MwAAAGYAMwBmAGYAZgCZAGYAzABmAP8A
ZgAAK2YAMytmAGYrZgCZK2YAzCtmAP8rZgAAVWYAM1VmAGZVZgCZVWYAzFVmAP9VZgAAgGYAM4BmAGaA
ZgCZgGYAzIBmAP+AZgAAqmYAM6pmAGaqZgCZqmYAzKpmAP+qZgAA1WYAM9VmAGbVZgCZ1WYAzNVmAP/V
ZgAA/2YAM/9mAGb/ZgCZ/2YAzP9mAP//ZgAAAJkAMwCZAGYAmQCZAJkAzACZAP8AmQAAK5kAMyuZAGYr
mQCZK5kAzCuZAP8rmQAAVZkAM1WZAGZVmQCZVZkAzFWZAP9VmQAAgJkAM4CZAGaAmQCZgJkAzICZAP+A
mQAAqpkAM6qZAGaqmQCZqpkAzKqZAP+qmQAA1ZkAM9WZAGbVmQCZ1ZkAzNWZAP/VmQAA/5kAM/+ZAGb/
mQCZ/5kAzP+ZAP//mQAAAMwAMwDMAGYAzACZAMwAzADMAP8AzAAAK8wAMyvMAGYrzACZK8wAzCvMAP8r
zAAAVcwAM1XMAGZVzACZVcwAzFXMAP9VzAAAgMwAM4DMAGaAzACZgMwAzIDMAP+AzAAAqswAM6rMAGaq
zACZqswAzKrMAP+qzAAA1cwAM9XMAGbVzACZ1cwAzNXMAP/VzAAA/8wAM//MAGb/zACZ/8wAzP/MAP//
zAAAAP8AMwD/AGYA/wCZAP8AzAD/AP8A/wAAK/8AMyv/AGYr/wCZK/8AzCv/AP8r/wAAVf8AM1X/AGZV
/wCZVf8AzFX/AP9V/wAAgP8AM4D/AGaA/wCZgP8AzID/AP+A/wAAqv8AM6r/AGaq/wCZqv8AzKr/AP+q
/wAA1f8AM9X/AGbV/wCZ1f8AzNX/AP/V/wAA//8AM///AGb//wCZ//8AzP//AP///wAAAAAAAAAAAAAA
AAAAAAAAHB0WHRwdHRwXHRwdHRwdHB0cFx0cHRwXHRwdHRwdHB0dFh0dHB0AAB0cHRwXHB0WHRwXHB0W
HRYdHB0cFxwXHB0WHRwXHBccHRwdFh0cAAAcHRYdHB0cHRwdHB0cHRwdHB0WHRwdHB0cHRwdHB0cHRwX
HB0cHQAAHRwdHBcdFh0cFxwdFh0XHB0dHB0WHRwXHBccFxwdFh0cHRwXHB0AAB0cFxwdHB0cHR0cHR0c
HRwdHBccHR0cHR0cHR0cHR0cHRccHRwdAAAdHB0dFh0cFxwXHBccFxwdFh0cHRYdFh0cFxwdFh0WHRwd
HBccHQAAHRwXHB0cHRwdHB0cHRwdHB0WHRwdHB0cHRwdHB0cHRwdFh0dHB0AAB0cHRwXHBccHRwXHB0c
Fx0cHRwXHB0cFxwXHRYdHBccHRwdFh0cAAAdHBccHRwdHBccHRwXHB0cHRYdHB0WHRwdHB0cHRwdHRwX
HB0cHQAAHRwdd/v7+/v7+/v7+/v7mh3R+/v7+/v1HBccHRYdHB0cHRwXHB0AAB0WHSL7+/v7+/v7+/v7
+/UcTfv7+/v7+3AdHB0dFh0WHRccHRwdAAAdHB0X0fv7+/v7+/v7+/v7mh3R+/v7+/uaHRYdHB0cHRwd
HBccHQAAFh0cHU37+/v7+/v7+/v7+/Udp/v7+/v7yxwdHB0cFxwdFh0dFh0AAB0cFxwd0fv7+/v7+/v7
+/v7cB37+/v7+/tAHRYdHB0dHB0cHRwdAAAdHB0cF6f7+/v7+8oXHB0cHR0c0fv7+/v7+/v7+/vEHRwd
Fh0cHQAAHRYdHRxN+/v7+/v1HB0XHB0WHXf7+/v7+/v7+/v79BccFxwdFxwAAB0cHRwXHNH7+/v7+3Ad
HB0cHRwd+/v7+/v7+/v7+/tHHB0cHRwdAAAdFh0cHR2n+/v7+/ubHB0WHRwdHdH7+/v7+/v7+/v7xB0c
HRYdHAAAHRwdHB0cTfv7+/v79B0cHRwdFh13+/v7+/v7+/v7+/UdFh0dHB0AAB0cHRccFx37+/v7+/tw
HRwXHB0cHfv7+/v7+0YdHB0cHRwcHRwXAAAcFxwdHB0c0fv7+/v7xRYdHRwdHB3R+/v7+/v7+/v7+/vF
HRYdHAAAHRwdFh0cHXf7+/v7+/tHHB0WHRccd/v7+/v7+/v7+/v79B0cHR0AABccHRwdFh0d+/v7+/v7
mhccHRwdHE37+/v7+/v7+/v7+/UdHBccAAAdHBcdHB0cHdH7+/v7+8scHRYdHB0d0fv7+/v7+/v7+/v7
cB0cHQAAHB0cHRwdFh13+/v7+/v7ah0dHB0WHaH7+/v7+/v7+/v7+8UcHR0AAB0cFxwXHB0cHRwdFh0c
HRwdHBccHRwdHB0WHRwXHB0cHRwXHBccAAAWHRwdHB0dHB0WHRwdHRwdFh0cHRwdHRYdHB0cHR0WHRYd
HB0cHQAAHRwXHB0cFxwXHB0WHRYdFh0cFxwXHBccHRYdFh0cHRwdHBccHRwAAB0cHR0WHRwdHRwdHRwd
HB0dHB0dHB0cHR0cHR0cHRwdFxwdFxwdAAAdFh0cHRwXHB0WHRwdFh0cFxwdHBccHRYdHB0WHRwXHB0c
HRwdHAAAHB0cHRYdHB0cHRwXHB0cHRwdFh0cHRwdHBccHRwdHB0WHRwdFh0AAB0cFxwdHBccHRYdHB0W
HRwXHB0cFxwdFh0cHRYdHBccHRwXHB0cAAAcHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0cHR0c
HR0cHQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
')



Function CheckForInstalledModules
{
    if ((get-module -ListAvailable "SkypeOnlineConnector").version.major -lt 7)	{
        [Microsoft.VisualBasic.Interaction]::MsgBox("Skype for Business PowerShell Module needs to be at least Version 7 as reported by Get-Module." ,'Information', "Please install the current Skype Online Connector.")
	exit
    }
    else {
	Import-Module SkypeOnlineConnector
	}
}

Function TestNumber()
{

    $TestNumberForm = New-Object System.Windows.Forms.Form 
    $TestNumberForm.Text = "Copy Normalization Rule to Dial Plan"
    $TestNumberForm.Size = New-Object System.Drawing.Size(400,240) 
    $TestNumberForm.MinimumSize = New-Object System.Drawing.Size(400,240) 
    $TestNumberForm.StartPosition = "CenterScreen"
    $TestNumberForm.KeyPreview = $True
    $TestNumberForm.Icon = $Global:TeamsAdminIcon

    $UserUPNLabel = New-Object System.Windows.Forms.Label
    $UserUPNLabel.Location = New-Object System.Drawing.Size(30,10) 
    $UserUPNLabel.Size = New-Object System.Drawing.Size(300,20) 
    $UserUPNLabel.Text = "Enter UPN of Dialing User:"
    $TestNumberForm.Controls.Add($UserUPNLabel)

    $UserUPNTextbox = New-Object System.Windows.Forms.Textbox
    $UserUPNTextbox.Location = New-Object System.Drawing.Size(30,30) 
    $UserUPNTextbox.Size = New-Object System.Drawing.Size(200,20) 
    $TestNumberForm.Controls.Add($UserUPNTextbox) 

    $DialedNumLabel = New-Object System.Windows.Forms.Label
    $DialedNumLabel.Location = New-Object System.Drawing.Size(30,55) 
    $DialedNumLabel.Size = New-Object System.Drawing.Size(300,20) 
    $DialedNumLabel.Text = "Enter Dialed Number:"
    $TestNumberForm.Controls.Add($DialedNumLabel)

    $DialedNumTextbox = New-Object System.Windows.Forms.Textbox
    $DialedNumTextbox.Location = New-Object System.Drawing.Size(30,75) 
    $DialedNumTextbox.Size = New-Object System.Drawing.Size(200,20) 
    $TestNumberForm.Controls.Add($DialedNumTextbox) 

    $DialedNumResultLabel = New-Object System.Windows.Forms.Label
    $DialedNumResultLabel.Location = New-Object System.Drawing.Size(30,100) 
    $DialedNumResultLabel.Size = New-Object System.Drawing.Size(300,20) 
    $TestNumberForm.Controls.Add($DialedNumResultLabel)

    $DialedNumResultNameLabel = New-Object System.Windows.Forms.Label
    $DialedNumResultNameLabel.Location = New-Object System.Drawing.Size(30,120) 
    $DialedNumResultNameLabel.Size = New-Object System.Drawing.Size(300,20) 
    $TestNumberForm.Controls.Add($DialedNumResultNameLabel)


    $TestNumButton = New-Object System.Windows.Forms.Button
    $TestNumButton.Location = New-Object System.Drawing.Size(30,160)
    $TestNumButton.Size = New-Object System.Drawing.Size(100,25)
    $TestNumButton.Text = "Test"
    $TestNumButton.Add_Click(
	{
    $DialedResult =Get-CsEffectiveTenantDialPlan -Identity $UserUPNTextBox.Text | Test-CsEffectiveTenantDialPlan -DialedNumber $DialedNumTextbox.Text
    $DialedNumResultLabel.Text = "Translation: $($DialedResult.TranslatedNumber)"
    $DialedNumResultNameLabel.Text = "Rule: $($DialedResult.MatchingRule.Name)"

	})
    $TestNumButton.Anchor = 'Bottom, Left'
    $TestNumberForm.Controls.Add($TestNumButton)

    $CancelTestNumButton = New-Object System.Windows.Forms.Button
    $CancelTestNumButton.Location = New-Object System.Drawing.Size(130,160)
    $CancelTestNumButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelTestNumButton.Text = "Quit"
    $CancelTestNumButton.Add_Click(
	{
	$TestNumberForm.Close()
	})
    $CancelTestNumButton.Anchor = 'Bottom, Left'
    $TestNumberForm.Controls.Add($CancelTestNumButton)

    [void] $TestNumberForm.ShowDialog()
}

Function ConnectToTenant()
{
    $connectForm = New-Object System.Windows.Forms.Form 
    $connectForm.Text = "Connect To Tenant"
    $connectForm.Size = New-Object System.Drawing.Size(300,200) 
    $connectForm.MinimumSize = New-Object System.Drawing.Size(300,200) 
    $connectForm.StartPosition = "CenterScreen"
    $connectForm.KeyPreview = $True
    $connectForm.Icon = $Global:TeamsAdminIcon

    $OnMicrosoftLabel = New-Object System.Windows.Forms.Label
    $OnMicrosoftLabel.Location = New-Object System.Drawing.Size(30,30) 
    $OnMicrosoftLabel.Size = New-Object System.Drawing.Size(300,20) 
    $OnMicrosoftLabel.Text = "Enter tenant's onmicrosoft.com domain"
    $connectForm.Controls.Add($OnMicrosoftLabel)

    $OnMicrosoftTextbox = New-Object System.Windows.Forms.Textbox
    $OnMicrosoftTextbox.Location = New-Object System.Drawing.Size(30,50) 
    $OnMicrosoftTextbox.Size = New-Object System.Drawing.Size(200,20) 
    
    $registryPath = "HKCU:\Software\TeamsScripts\Scripts"
    $Name = "OnMicrosoftDomain"
    IF((Test-Path $registryPath))
	{
	$OnMicrosoftTextbox.text = (Get-ItemProperty -Path HKCU:\Software\TeamsScripts\Scripts -Name OnMicrosoftDomain).OnMicrosoftDomain
	}
    $connectForm.Controls.Add($OnMicrosoftTextbox) 

    $SaveDomainLabel = New-Object System.Windows.Forms.Label
    $SaveDomainLabel.Location = New-Object System.Drawing.Size(30,80) 
    $SaveDomainLabel.Size = New-Object System.Drawing.Size(170,20) 
    $SaveDomainLabel.Text = "Save tenant domain in registry?"
    $connectForm.Controls.Add($SaveDomainLabel)

    $SaveDomainCheckbox = New-Object System.Windows.Forms.Checkbox
    $SaveDomainCheckbox.Location = New-Object System.Drawing.Size(215,80) 
    $SaveDomainCheckbox.Size = New-Object System.Drawing.Size(20,20) 
    $connectForm.Controls.Add($SaveDomainCheckbox) 

    $AcceptConnectButton = New-Object System.Windows.Forms.Button
    $AcceptConnectButton.Location = New-Object System.Drawing.Size(30,110)
    $AcceptConnectButton.Size = New-Object System.Drawing.Size(100,25)
    $AcceptConnectButton.Text = "Connect"
    $AcceptConnectButton.Add_Click(
	{
	#if box is checked, save that key
	if ($SaveDomainCheckbox.checked -eq $true)
		{
		#If the key doesn't exist
		IF(!(Test-Path $registryPath))
			{
			New-Item -Path $registryPath -Force | Out-Null
			New-ItemProperty -Path $registryPath -Name $name -Value $OnMicrosoftTextbox.text -PropertyType String -Force | Out-Null
			}
		ELSE 
			{
			New-ItemProperty -Path $registryPath -Name $name -Value $OnMicrosoftTextbox.text -PropertyType String -Force | Out-Null
			}
		}

	$AcceptConnectButton.Text = "Connecting"
	$AcceptConnectButton.enabled=$false
	$global:sfbSession = New-CsOnlineSession -overrideadmindomain $OnMicrosofttextbox.text
	Import-PSSession $global:sfbSession -allowclobber
	$global:dialplans=get-cstenantdialplan
        $TeamsListBox.Items.Clear()
	    foreach ($plan in $dialplans) 
	    {
		    [void] $TeamsListBox.Items.Add($plan.simplename)
	    }

	$AcceptConnectButton.enabled=$true
	$AcceptConnectButton.Text = "Connect"
	$connectForm.Close()
	$global:connected=$true
	})
    $AcceptConnectButton.Anchor = 'Bottom, Left'
    $connectForm.Controls.Add($AcceptConnectButton)

    $CancelConnectButton = New-Object System.Windows.Forms.Button
    $CancelConnectButton.Location = New-Object System.Drawing.Size(130,110)
    $CancelConnectButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelConnectButton.Text = "Quit"
    $CancelConnectButton.Add_Click(
	{
	$connectForm.Close()
	})
    $CancelConnectButton.Anchor = 'Bottom, Left'
    $connectForm.Controls.Add($CancelConnectButton)

    [void] $connectForm.ShowDialog()
}


Function CopyToDP($nrname)
{
    $CopyToDPForm = New-Object System.Windows.Forms.Form 
    $CopyToDPForm.Text = "Copy Normalization Rule to Dial Plan"
    $CopyToDPForm.Size = New-Object System.Drawing.Size(300,140) 
    $CopyToDPForm.MinimumSize = New-Object System.Drawing.Size(300,140) 
    $CopyToDPForm.StartPosition = "CenterScreen"
    $CopyToDPForm.KeyPreview = $True
    $CopyToDPForm.Icon = $Global:TeamsAdminIcon

    $ShowNRNameLabel = New-Object System.Windows.Forms.Label
    $ShowNRNameLabel.Location = New-Object System.Drawing.Size(30,10) 
    $ShowNRNameLabel.Size = New-Object System.Drawing.Size(300,20) 
    $ShowNRNameLabel.Text = "Copy Rule $NRName to:"
    $CopyToDPForm.Controls.Add($ShowNRNameLabel)

    $OnMicrosoftDropdown = New-Object System.Windows.Forms.ComboBox
    $OnMicrosoftDropdown.Location = New-Object System.Drawing.Size(30,30) 
    $OnMicrosoftDropdown.Size = New-Object System.Drawing.Size(200,20) 
    $OnMicrosoftDropdown.Items.Clear()
	    foreach ($plan in $dialplans) 
	    {
		    [void] $OnMicrosoftDropdown.Items.Add($plan.simplename)
	    }
    $CopyToDPForm.Controls.Add($OnMicrosoftDropdown) 

    $AcceptCopyButton = New-Object System.Windows.Forms.Button
    $AcceptCopyButton.Location = New-Object System.Drawing.Size(30,60)
    $AcceptCopyButton.Size = New-Object System.Drawing.Size(100,25)
    $AcceptCopyButton.Text = "Copy"
    $AcceptCopyButton.Add_Click(
	{
	$global:selectedDPforCopy= $OnMicrosoftDropDown.SelectedItem.ToString()
	$CopyToDPForm.Close()
	})
    $AcceptCopyButton.Anchor = 'Bottom, Left'
    $CopyToDPForm.Controls.Add($AcceptCopyButton)

    $CancelCopyButton = New-Object System.Windows.Forms.Button
    $CancelCopyButton.Location = New-Object System.Drawing.Size(130,60)
    $CancelCopyButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelCopyButton.Text = "Quit"
    $CancelCopyButton.Add_Click(
	{
	$CopyToDPForm.Close()
	})
    $CancelCopyButton.Anchor = 'Bottom, Left'
    $CopyToDPForm.Controls.Add($CancelCopyButton)

    [void] $CopyToDPForm.ShowDialog()
}

Function MainForm()
{
	
    $mainForm = New-Object System.Windows.Forms.Form 
    $mainForm.Text = "Microsoft Teams Dial Plan Editor v 1.0.0"
    $mainForm.Size = New-Object System.Drawing.Size(980,560) 
    $mainForm.MinimumSize = New-Object System.Drawing.Size(980,560) 
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.Add_SizeChanged($CAC_FormSizeChanged)
    $mainForm.KeyPreview = $True
    $mainForm.Icon = $Global:TeamsAdminIcon

    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Location = New-Object System.Drawing.Size(10,10) 
    $TitleLabel.Size = New-Object System.Drawing.Size(780,30) 
    $TitleLabel.Text = "The purpose of this tool is to give you an easy graphical method of editing your dial plans in your tenant without having to remember a bunch of PowerShell commands you'd rarely use.  This is an early version with more features coming soon."
    $mainForm.Controls.Add($TitleLabel) 

    $DialPlansListBoxLabel = New-Object System.Windows.Forms.Label
    $DialPlansListBoxLabel.Location = New-Object System.Drawing.Size(10,65) 
    $DialPlansListBoxLabel.Size = New-Object System.Drawing.Size(250,15) 
    $DialPlansListBoxLabel.Text = "Dial Plans"
    $mainForm.Controls.Add($DialPlansListBoxLabel) 

    $TeamsListBox = New-Object System.Windows.Forms.ListBox 
    $TeamsListBox.Location = New-Object System.Drawing.Size(10,80) 
    $TeamsListBox.Size = New-Object System.Drawing.Size(250,370) 
    $TeamsListBox.Anchor = 'Top, Bottom,Left'
    $TeamsListBox.Sorted = $True
    $TeamsListBox.SelectionMode = "MultiExtended"
    $TeamsListBox.add_SelectedIndexChanged({
	$DialPlanDatagrid.Rows.Clear()
        $SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
        foreach ($NormRule in $SelectedDialPlan.normalizationrules) { $DialPlanDatagrid.Rows.Add($normrule.name,$normrule.pattern,$normrule.translation,$normrule.description,$normrule.isextension)}
    })
    $mainForm.Controls.Add($TeamsListBox) 

    $DialPlanDataGridLabel = New-Object System.Windows.Forms.Label
    $DialPlanDataGridLabel.Location = New-Object System.Drawing.Size(270,65) 
    $DialPlanDataGridLabel.Size = New-Object System.Drawing.Size(250,15) 
    $DialPlanDataGridLabel.Text = "Normalization Rules"
    $mainForm.Controls.Add($DialPlanDataGridLabel) 
	
    $DialPlanDatagrid = New-Object System.Windows.Forms.DataGridView
    $DialPlanDatagrid.Location = New-Object System.Drawing.Size(270,80) 
    $DialPlanDatagrid.Size = New-Object System.Drawing.Size(345,365) 
    $DialPlanDatagrid.Anchor = 'Top, Bottom,Left'
    $DialPlanDatagrid.RowsDefaultCellStyle.BackColor = [System.Drawing.Color]::Bisque 
    $DialPlanDatagrid.AlternatingRowsDefaultCellStyle.BackColor = [System.Drawing.Color]::Beige
    $DialPlanDatagrid.ColumnCount = 3
	#$DialPlanDatagrid.Columns[0].Width = 245
	#$dataGridView.Columns[1].Width = 245
	$DialPlanDatagrid.Columns[0].Name = "Name"
	$DialPlanDatagrid.Columns[1].Name = "Pattern"
	$DialPlanDatagrid.Columns[2].Name = "Translation"

    $DialPlanDatagrid.Add_Click(
	{
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	$NRPatternTextbox.Text = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].pattern
	$NRDescriptionTextbox.Text = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].description
	$NRTranslationTextbox.Text = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].translation
	$NRNameTextbox.Text = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].name
        if($SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].IsInternalExtension -eq "True") {$NRExtensionCheckbox.Checked = $true} else {$NRExtensionCheckbox.Checked = $false}
	})
    $mainForm.Controls.Add($DialPlanDatagrid) 

    $NRDescriptionLabel = New-Object System.Windows.Forms.Label
    $NRDescriptionLabel.Location = New-Object System.Drawing.Size(625,80) 
    $NRDescriptionLabel.Size = New-Object System.Drawing.Size(210,20) 
    $NRDescriptionLabel.Text = "Description"
    $mainForm.Controls.Add($NRDescriptionLabel) 

    $NRDescriptionTextbox = New-Object System.Windows.Forms.Textbox
    $NRDescriptionTextbox.Location = New-Object System.Drawing.Size(625,100) 
    $NRDescriptionTextbox.Size = New-Object System.Drawing.Size(210,20) 
    $NRDescriptionTextbox.Text = ""
    $mainForm.Controls.Add($NRDescriptionTextbox) 

    $NRPatternLabel = New-Object System.Windows.Forms.Label
    $NRPatternLabel.Location = New-Object System.Drawing.Size(625,130) 
    $NRPatternLabel.Size = New-Object System.Drawing.Size(210,20) 
    $NRPatternLabel.Text = "Pattern"
    $mainForm.Controls.Add($NRPatternLabel) 

    $NRPatternTextbox = New-Object System.Windows.Forms.Textbox
    $NRPatternTextbox.Location = New-Object System.Drawing.Size(625,150) 
    $NRPatternTextbox.Size = New-Object System.Drawing.Size(210,20) 
    $NRPatternTextbox.Text = ""
    $mainForm.Controls.Add($NRPatternTextbox) 

    $NRTranslationLabel = New-Object System.Windows.Forms.Label
    $NRTranslationLabel.Location = New-Object System.Drawing.Size(625,180) 
    $NRTranslationLabel.Size = New-Object System.Drawing.Size(210,20) 
    $NRTranslationLabel.Text = "Translation"
    $mainForm.Controls.Add($NRTranslationLabel) 

    $NRTranslationTextbox = New-Object System.Windows.Forms.Textbox
    $NRTranslationTextbox.Location = New-Object System.Drawing.Size(625,200) 
    $NRTranslationTextbox.Size = New-Object System.Drawing.Size(210,20) 
    $NRTranslationTextbox.Text = ""
    $mainForm.Controls.Add($NRTranslationTextbox) 

    $NRNameLabel = New-Object System.Windows.Forms.Label
    $NRNameLabel.Location = New-Object System.Drawing.Size(625,230) 
    $NRNameLabel.Size = New-Object System.Drawing.Size(210,20) 
    $NRNameLabel.Text = "Name"
    $mainForm.Controls.Add($NRNameLabel) 

    $NRNameTextbox = New-Object System.Windows.Forms.Textbox
    $NRNameTextbox.Location = New-Object System.Drawing.Size(625,250) 
    $NRNameTextbox.Size = New-Object System.Drawing.Size(210,20) 
    $NRNameTextbox.Text = ""
    $mainForm.Controls.Add($NRNameTextbox) 

    $NRExtensionLabel = New-Object System.Windows.Forms.Label
    $NRExtensionLabel.Location = New-Object System.Drawing.Size(625,280) 
    $NRExtensionLabel.Size = New-Object System.Drawing.Size(210,20) 
    $NRExtensionLabel.Text = "Is Internal Extension?"
    $mainForm.Controls.Add($NRExtensionLabel) 

    $NRExtensionCheckbox = New-Object System.Windows.Forms.Checkbox
    $NRExtensionCheckbox.Location = New-Object System.Drawing.Size(625,300) 
    $NRExtensionCheckbox.Size = New-Object System.Drawing.Size(210,20) 
    $mainForm.Controls.Add($NRExtensionCheckbox) 


    $SaveEditButton = New-Object System.Windows.Forms.Button
    $SaveEditButton.Location = New-Object System.Drawing.Size(625, 330)
    $SaveEditButton.Size = New-Object System.Drawing.Size(100,25)
    $SaveEditButton.Text = "Save Edit"
    $SaveEditButton.Add_Click({
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	$NRName = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].name
	$NRDelete = (Get-CsTenantDialPlan $SelectedDialPlan.identity).normalizationrules|where {$_.name -like $NRName} 
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{remove=$NRDelete}
	$NR = @()
	$NR += New-CsVoiceNormalizationRule -Name $NRNameTextbox.Text -Parent $SelectedDialPlan.identity -Pattern $NRPatterntextbox.Text -Translation $NRTranslationtextbox.Text -Description $NRDescriptiontextbox.Text -IsInternalExtension:$NRExtensionCheckbox.Checked -Priority 0 -InMemory
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{add=$NR}
	$global:dialplans=get-cstenantdialplan
	$DialPlanDatagrid.Rows.Clear()
        $SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	foreach ($NormRule in $SelectedDialPlan.normalizationrules) { $DialPlanDatagrid.Rows.Add($normrule.name,$normrule.pattern,$normrule.translation,$normrule.description,$normrule.isextension)}
    })
    $SaveEditButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($SaveEditButton)

    $DeleteRuleButton = New-Object System.Windows.Forms.Button
    $DeleteRuleButton.Location = New-Object System.Drawing.Size(730,330)
    $DeleteRuleButton.Size = New-Object System.Drawing.Size(100,25)
    $DeleteRuleButton.Text = "Delete Rule"
    $DeleteRuleButton.Add_Click({
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	$NRName = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].name
	$NRDelete = (Get-CsTenantDialPlan $SelectedDialPlan.identity).normalizationrules|where {$_.name -like $NRName} 
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{remove=$NRDelete}
	$global:dialplans=get-cstenantdialplan
	$DialPlanDatagrid.Rows.Clear()
        $SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	foreach ($NormRule in $SelectedDialPlan.normalizationrules) { $DialPlanDatagrid.Rows.Add($normrule.name,$normrule.pattern,$normrule.translation,$normrule.description,$normrule.isextension)}
    })
    $DeleteRuleButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($DeleteRuleButton)

    $AddNewNRButton = New-Object System.Windows.Forms.Button
    $AddNewNRButton.Location = New-Object System.Drawing.Size(625,360)
    $AddNewNRButton.Size = New-Object System.Drawing.Size(100,25)
    $AddNewNRButton.Text = "Add New"
    $AddNewNRButton.Add_Click({
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	$NR = @()
	$NR += New-CsVoiceNormalizationRule -Name $NRNameTextbox.Text -Parent $SelectedDialPlan.identity -Pattern $NRPatterntextbox.Text -Translation $NRTranslationtextbox.Text -Description $NRDescriptiontextbox.Text -IsInternalExtension:$NRExtensionCheckbox.Checked -Priority 0 -InMemory
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{add=$NR}
	$global:dialplans=get-cstenantdialplan
	$DialPlanDatagrid.Rows.Clear()
        $SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	foreach ($NormRule in $SelectedDialPlan.normalizationrules) { $DialPlanDatagrid.Rows.Add($normrule.name,$normrule.pattern,$normrule.translation,$normrule.description,$normrule.isextension)}
 
   })
    $AddNewNRButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($AddNewNRButton)

    $CancelEditButton = New-Object System.Windows.Forms.Button
    $CancelEditButton.Location = New-Object System.Drawing.Size(730,360)
    $CancelEditButton.Size = New-Object System.Drawing.Size(100,25)
    $CancelEditButton.Text = "Cancel Edit"
    $CancelEditButton.Add_Click({
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	$NRPatternTextbox.Text = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].pattern
	$NRDescriptionTextbox.Text = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].description
	$NRTranslationTextbox.Text = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].translation
	$NRNameTextbox.Text = $SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].name
        if($SelectedDialPlan.normalizationrules[$DialPlanDatagrid.CurrentCell.RowIndex].IsInternalExtension -eq "True") {$NRExtensionCheckbox.Checked = $true} else {$NRExtensionCheckbox.Checked = $false}


    })
    $CancelEditButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($CancelEditButton)

    $DuplicateNRButton = New-Object System.Windows.Forms.Button
    $DuplicateNRButton.Location = New-Object System.Drawing.Size(625,390)
    $DuplicateNRButton.Size = New-Object System.Drawing.Size(100,25)
    $DuplicateNRButton.Text = "Duplicate"
    $DuplicateNRButton.Add_Click({
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	$NR = @()
	$NR += New-CsVoiceNormalizationRule -Name $($NRNameTextbox.Text + "-Copy") -Parent $SelectedDialPlan.identity -Pattern $NRPatterntextbox.Text -Translation $NRTranslationtextbox.Text -Description $NRDescriptiontextbox.Text -IsInternalExtension:$NRExtensionCheckbox.Checked -Priority 0 -InMemory
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{add=$NR}
	$global:dialplans=get-cstenantdialplan
	$DialPlanDatagrid.Rows.Clear()
        $SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	foreach ($NormRule in $SelectedDialPlan.normalizationrules) { $DialPlanDatagrid.Rows.Add($normrule.name,$normrule.pattern,$normrule.translation,$normrule.description,$normrule.isextension)}
 
   })
    $DuplicateNRButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($DuplicateNRButton)

    $CopyToDPButton = New-Object System.Windows.Forms.Button
    $CopyToDPButton.Location = New-Object System.Drawing.Size(730,390)
    $CopyToDPButton.Size = New-Object System.Drawing.Size(100,25)
    $CopyToDPButton.Text = "Copy To DP"
    $CopyToDPButton.Add_Click({
	$SelectedDialPlan=CopyToDP $NRNameTextbox.Text
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $global:selectedDPforCopy}
	$NR = @()
	$NR += New-CsVoiceNormalizationRule -Name $NRNameTextbox.Text -Parent $SelectedDialPlan.identity -Pattern $NRPatterntextbox.Text -Translation $NRTranslationtextbox.Text -Description $NRDescriptiontextbox.Text -IsInternalExtension:$NRExtensionCheckbox.Checked -Priority 0 -InMemory
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{add=$NR}
	$global:dialplans=get-cstenantdialplan
	$DialPlanDatagrid.Rows.Clear()
        $SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	foreach ($NormRule in $SelectedDialPlan.normalizationrules) { $DialPlanDatagrid.Rows.Add($normrule.name,$normrule.pattern,$normrule.translation,$normrule.description,$normrule.isextension)}

    })
    $CopyToDPButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($CopyToDPButton)

    $MoveUpButton = New-Object System.Windows.Forms.Button
    $MoveUpButton.Location = New-Object System.Drawing.Size(625,420)
    $MoveUpButton.Size = New-Object System.Drawing.Size(100,25)
    $MoveUpButton.Text = "Move Up"
    $MoveUpButton.Add_Click({
	$MoveUpButton.enabled=$false
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	$NR = $SelectedDialPlan.normalizationrules
	$NewNR=@()
	if ($DialPlanDatagrid.CurrentCell.RowIndex -gt 0)
		{
        $currentindex=$DialPlanDatagrid.CurrentCell.RowIndex
		for ($i=0;$i -lt ($DialPlanDatagrid.CurrentCell.RowIndex -1) ; $i++) {$NewNR += $NR[$i]}
		$NewNR += $NR[$DialPlanDatagrid.CurrentCell.RowIndex]
		$NewNR += $NR[$DialPlanDatagrid.CurrentCell.RowIndex-1]
		for ($i=$DialPlanDatagrid.CurrentCell.RowIndex+1;$i -lt ($DialPlanDatagrid.Rows.count - 1) ; $i++) {$NewNR += $NR[$i]}
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{remove=$NR}
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{add=$NewNR}
	$global:dialplans=get-cstenantdialplan
	$DialPlanDatagrid.Rows.Clear()
        $SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	foreach ($NormRule in $SelectedDialPlan.normalizationrules) { $DialPlanDatagrid.Rows.Add($normrule.name,$normrule.pattern,$normrule.translation,$normrule.description,$normrule.isextension)}
    #CurrentCell.RowIndex = 
    #$DialPlanDatagrid.Rows[$currentindex-1].Selected = $true
    $DialPlanDatagrid.CurrentCell = $DialPlanDatagrid.Rows[$currentindex - 1].Cells[0]
		}
	$moveupbutton.enabled=$true
    })
    $MoveUpButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($MoveUpButton)


    $MoveDownButton = New-Object System.Windows.Forms.Button
    $MoveDownButton.Location = New-Object System.Drawing.Size(730,420)
    $MoveDownButton.Size = New-Object System.Drawing.Size(100,25)
    $MoveDownButton.Text = "Move Down"
    $MoveDownButton.Add_Click({
	$MoveDownButton.Enabled=$false
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	$NR = $SelectedDialPlan.normalizationrules
	$NewNR=@()
	if ($DialPlanDatagrid.CurrentCell.RowIndex -lt ($DialPlanDatagrid.Rows.count - 2))
		{
        $currentindex=$DialPlanDatagrid.CurrentCell.RowIndex
		for ($i=0;$i -lt $DialPlanDatagrid.CurrentCell.RowIndex ; $i++) {$NewNR += $NR[$i]}
		$NewNR += $NR[$DialPlanDatagrid.CurrentCell.RowIndex+1]
		$NewNR += $NR[$DialPlanDatagrid.CurrentCell.RowIndex]
		for ($i=$DialPlanDatagrid.CurrentCell.RowIndex+2;$i -lt ($DialPlanDatagrid.Rows.count - 1) ; $i++) {$NewNR += $NR[$i]}
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{remove=$NR}
	Set-CsTenantDialPlan -Identity $SelectedDialPlan.identity -NormalizationRules @{add=$NewNR}
	$global:dialplans=get-cstenantdialplan
	$DialPlanDatagrid.Rows.Clear()
        $SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	foreach ($NormRule in $SelectedDialPlan.normalizationrules) { $DialPlanDatagrid.Rows.Add($normrule.name,$normrule.pattern,$normrule.translation,$normrule.description,$normrule.isextension)}

		}
    $DialPlanDatagrid.CurrentCell = $DialPlanDatagrid.Rows[$currentindex + 1].Cells[0]
	$MoveDownButton.Enabled=$True
    })
    $MoveDownButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($MoveDownButton)

    $ConnectTenantButton = New-Object System.Windows.Forms.Button
    $ConnectTenantButton.Location = New-Object System.Drawing.Size((10 + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 100))
    $ConnectTenantButton.Size = New-Object System.Drawing.Size(115,35)
    $ConnectTenantButton.Text = "Connect to Tenant"
    $ConnectTenantButton.Add_Click({
	ConnectToTenant

    })
    $ConnectTenantButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($ConnectTenantButton)

 
    $RefreshDPButton = New-Object System.Windows.Forms.Button
    $RefreshDPButton.Location = New-Object System.Drawing.Size((135 + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 100))
    $RefreshDPButton.Size = New-Object System.Drawing.Size(115,35)
    $RefreshDPButton.Text = "Refresh Dial Plans"
    $RefreshDPButton.Add_Click({
	$global:dialplans=get-cstenantdialplan
        $TeamsListBox.Items.Clear()
	    foreach ($plan in $global:dialplans) 
	    {
		    [void] $TeamsListBox.Items.Add($plan.simplename)
	    }

    })
    $RefreshDPButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($RefreshDPButton)

    $AddNewDPButton = New-Object System.Windows.Forms.Button
    $AddNewDPButton.Location = New-Object System.Drawing.Size((260 + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 100))
    $AddNewDPButton.Size = New-Object System.Drawing.Size(115,35)
    $AddNewDPButton.Text = "Create New Dial Plan"
    $AddNewDPButton.Add_Click({
	#Check if connected to tenant
	if ($global:connected -eq $true)
	{
	$newDPName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the name of the new dial plan.", "New Dial Plan")
	New-CSTenantDialPlan -identity $newDPName
	$global:dialplans=get-cstenantdialplan
        $TeamsListBox.Items.Clear()
	    foreach ($plan in $global:dialplans) 
	    {
		    [void] $TeamsListBox.Items.Add($plan.simplename)
	    }
	}	

    })
    $AddNewDPButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($AddNewDPButton)

    $CopyDPButton = New-Object System.Windows.Forms.Button
    $CopyDPButton.Location = New-Object System.Drawing.Size((385 + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 100))
    $CopyDPButton.Size = New-Object System.Drawing.Size(115,35)
    $CopyDPButton.Text = "Copy Dial Plan"
    $CopyDPButton.Add_Click({
	if ($global:connected -eq $true)
	{
	$SelectedDialPlan=$global:dialplans |where {$_.simplename -eq $TeamsListBox.SelectedItem.tostring()}
	$newDPName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter name of new dial plan where $($SelectedDialplan.simplename) will be copied.", "Copy Dial Plan")
	if ($selectedDialplan.description -eq $null) { $selectedDialplan.description=""}
	if ($SelectedDialPlan.ExternalAccessPrefix -eq $null) { New-CSTenantDialPlan -identity $newDPName -description $selectedDialplan.description -normalizationrules $SelectedDialPlan.NormalizationRules -OptimizeDeviceDialing $SelectedDialPlan.OptimizeDeviceDialing}
	else { New-CSTenantDialPlan -identity $newDPName -description $selectedDialplan.description -normalizationrules $SelectedDialPlan.NormalizationRules -ExternalAccessPrefix $SelectedDialPlan.ExternalAccessPrefix -OptimizeDeviceDialing $SelectedDialPlan.OptimizeDeviceDialing}
	$global:dialplans=get-cstenantdialplan
        $TeamsListBox.Items.Clear()
	    foreach ($plan in $global:dialplans) 
	    {
		    [void] $TeamsListBox.Items.Add($plan.simplename)
	    }
	}	
    })
    $CopyDPButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($CopyDPButton)

 
    $DeleteDPButton = New-Object System.Windows.Forms.Button
    $DeleteDPButton.Location = New-Object System.Drawing.Size((510 + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 100))
    $DeleteDPButton.Size = New-Object System.Drawing.Size(115,35)
    $DeleteDPButton.Text = "Delete Dial Plan"
    $DeleteDPButton.Add_Click({
	#Check if connected to tenant
	if ($global:connected -eq $true)
	{
	
        $deleteresponse=[Microsoft.VisualBasic.Interaction]::MsgBox("Are you sure you want to delete $($TeamsListBox.SelectedItem.tostring())?  This deletion can not be undone!" ,'YesNo', "Confirm Dial Plan Deletion")
	if ($deleteresponse -eq "Yes")	
		{
		$dptoremove=$global:dialplans | where {$_.simplename -like $TeamsListBox.SelectedItem.tostring()}
		remove-cstenantdialplan $dptoremove.identity
		$global:dialplans=get-cstenantdialplan
        	$TeamsListBox.Items.Clear()
	    	foreach ($plan in $global:dialplans) 
	    		{
		    	[void] $TeamsListBox.Items.Add($plan.simplename)
	 		}
		}
	}	

    })
    $DeleteDPButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($DeleteDPButton)

    $RunTestButton = New-Object System.Windows.Forms.Button
    $RunTestButton.Location = New-Object System.Drawing.Size((635 + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 100))
    $RunTestButton.Size = New-Object System.Drawing.Size(115,35)
    $RunTestButton.Text = "Test Number"
    $RunTestButton.Add_Click(
	{
	TestNumber
	})
    $RunTestButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($RunTestButton)


    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size((760 + (($mainForm.width-50) /7 * 0) ),($mainForm.height - 100))
    $CancelButton.Size = New-Object System.Drawing.Size(115,35)
    $CancelButton.Text = "Quit"
    $CancelButton.Add_Click(
	{
	$mainForm.Close()
	remove-pssession -session $global:sfbSession
	})
    $CancelButton.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($CancelButton)

    #TeamsAdmin LinkLabel
    $TeamsAdminLinkLabel = New-Object System.Windows.Forms.LinkLabel
    $TeamsAdminLinkLabel.Location = New-Object System.Drawing.Size(10,($mainForm.height - 60)) 
    $TeamsAdminLinkLabel.Size = New-Object System.Drawing.Size(200,20)
    $TeamsAdminLinkLabel.text = "http://www.TeamsAdmin.com"
    $TeamsAdminLinkLabel.add_Click({Start-Process $TeamsAdminLinkLabel.text})
    $TeamsAdminLinkLabel.Anchor = 'Bottom, Left'
    $mainForm.Controls.Add($TeamsAdminLinkLabel)

    [void] $mainForm.ShowDialog()

}

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$global:connected=$false

CheckForInstalledModules
MainForm
