# Employment Tool v3
#	Tools - Phone Updater

$adList = Get-ADUser -Filter * -SearchBase "CN=Users,DC=domain,DC=com"

	# Prompt for username and return existing information

Function PhoneUpdater {
$form_PhoneUpdater = New-Form
$form_PhoneUpdater.Size = Set-Size 300 200
$form_PhoneUpdater.Text = "  Employment Tool - Phone Updater"
$form_PhoneUpdater.SizeGripStyle = 'Hide'
$form_PhoneUpdater.StartPosition = 'CenterScreen'
$form_PhoneUpdater.FormBorderStyle = 'Fixed3D'
$form_PhoneUpdater.Font = "Arial,10,style=Bold"
$form_PhoneUpdater.TopMost = $True
$form_PhoneUpdater.KeyPreview = $True
$form_PhoneUpdater.MaximizeBox = $False
$form_PhoneUpdater.MinimizeBox = $False
$form_PhoneUpdater.Icon = Set-Icon

$label_Username = New-Label
$label_Username.Size = Set-Size 100 20
$label_Username.Location = Set-Point 10 10
$label_Username.Text = "Username"
	$form_PhoneUpdater.Controls.Add($label_Username)

$textbox_Username = New-Textbox
$textbox_Username.Size = Set-Size 150 20
$textbox_Username.Location = Set-Point 120 10
$textbox_Username.Add_TextChanged({
	$checkAD = $adList.SamAccountName |Select-String "^$($textbox_Username.Text)$" -Quiet
	If ($checkAD -eq $True) {
		$label_Username.Forecolor = "Green"
		$button_Submit.Enabled = $True
		$script:theUser = $textbox_Username.Text
		$userInfo = Get-ADUser $theUser -Properties *
		$script:currentExt = "$($userInfo.ipPhone)"
		$textbox_Ext.Text = $currentExt
		$script:currentDID = "$($userInfo.Pager)"
		$textbox_DID.Text = $currentDID
		$script:currentMobile = "$($userInfo.Mobile)"
		$textbox_Mobile.Text = $currentMobile
		} Else {
			Clear-Form
			}
	})
	$form_PhoneUpdater.Controls.Add($textbox_Username)

$label_Ext = New-Label
$label_Ext.Size = Set-Size 100 20
$label_Ext.Location = Set-Point 10 40
$label_Ext.Text = "Extension"
	$form_PhoneUpdater.Controls.Add($label_Ext)

$textbox_Ext = New-Textbox
$textbox_Ext.Size = Set-Size 150 20
$textbox_Ext.Location = Set-Point 120 40
$textbox_Ext.MaxLength = 4
$textbox_Ext.Add_TextChanged({
	If ($textbox_Ext.Text -ne "") {
	$string = $textbox_Ext.Text 
		If ($textbox_Ext.Text -cmatch "^[0-9]*$") {} Else {
			Throw-Error "Please only use numbers."
			If ($currentExt) {$textbox_Ext.Text = $currentExt} Else {
				$textbox_Ext.Text = ""
				}
			}	
		}
	})
	$form_PhoneUpdater.Controls.Add($textbox_Ext)

$label_DID = New-Label
$label_DID.Size = Set-Size 100 20
$label_DID.Location = Set-Point 10 70
$label_DID.Text = "Direct Dial"
	$form_PhoneUpdater.Controls.Add($label_DID)

$textbox_DID = New-Textbox
$textbox_DID.Size = Set-Size 150 20
$textbox_DID.Location = Set-Point 120 70
$textbox_DID.MaxLength = 10
$textbox_DID.Add_TextChanged({
	If ($textbox_DID.Text -ne "") {
	$string = $textbox_DID.Text
		If ($textbox_DID.Text -cmatch "^[0-9]*$") {} Else {
			Throw-Error "Please use only numbers."
			If ($currentDID) {$textbox_DID.Text = $currentDID} Else {
				$textbox_DID.Text = ""
				}
			}
		}
	})
	$form_PhoneUpdater.Controls.Add($textbox_DID)
	
$label_Mobile = New-Label
$label_Mobile.Size = Set-Size 100 20
$label_Mobile.Location = Set-Point 10 100
$label_Mobile.Text = "Mobile"
	$form_PhoneUpdater.Controls.Add($label_Mobile)

$textbox_Mobile = New-Textbox
$textbox_Mobile.Size = Set-Size 150 20
$textbox_Mobile.Location = Set-Point 120 100
$textbox_Mobile.MaxLength = 10
$textbox_Mobile.Add_TextChanged({
	If ($textbox_Mobile.Text -ne "") {
	$string = $textbox_Mobile.Text
		If ($textbox_Mobile.Text -cmatch "^[0-9]*$") {} Else {
			Throw-Error "Please use only numbers."
			If ($currentMobile) {$textbox_Mobile.Text = $currentMobile} Else {
				$textbox_Mobile.Text = ""
				}
			}
		}
	})
	$form_PhoneUpdater.Controls.Add($textbox_Mobile)

$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 140 130
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_PhoneUpdater.Controls.Add($button_Cancel)
	$form_PhoneUpdater.CancelButton = $button_Cancel
	
$button_Submit = New-Button
$button_Submit.Size = Set-Size 60 30
$button_Submit.Location = Set-Point 210 130
$button_Submit.Text = "Submit"
$button_Submit.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_PhoneUpdater.Controls.Add($button_Submit)
	$form_PhoneUpdater.AcceptButton = $button_Submit

$fr = $form_PhoneUpdater.ShowDialog()
If ($fr -eq "Cancel") { $form_PhoneUpdater.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") {
# Check to make sure fields are filled out as expected
# Ext = 0 or 4
# DID = 0 or 10
# Mobile = 0 or 10
# End check

	$newExt = $textbox_Ext.Text
	$newDID = $textbox_DID.Text
	$newMobile = $textbox_Mobile.Text
	
# EXT
	If ($currentExt -ne $newExt -AND $newExt -ne "") {Set-ADUser $theUser -Replace @{ipPhone="$newExt"}} # Set to new ext
	If ($currentExt -ne $newExt -AND $newExt -eq "") {Set-ADUser $theUser -Clear ipPhone} # Set ext to blank
# DID
	If ($currentDID -ne $newDID) {Set-ADUser $theUser -Replace @{pager="$newDID"}} # Set to new DID
	If ($currentDID -ne $newDID -AND $newDID -eq "") {Set-ADUser $theUser -Clear pager} # Set DID to blank
# MOBILE
	If ($currentMobile -ne $newMobile -AND $newMobile -ne "") {Set-ADUser $theUser -Replace @{mobile="$newMobile"}} # Set to new mobile
	If ($currentMobile -ne $newMobile -AND $newMobile -eq "") {Set-ADUser $theUser -Clear mobile} # Set mobile to blank
	
	$form_PhoneUpdater.Dispose()
	Run-MainMenu
	}
}

	# Reset the form

Function Clear-Form {
$label_Username.Forecolor = "Red"
$button_Submit.Enabled = $False
Remove-Variable theUser
Remove-Variable currentDID
Remove-Variable currentExt
Remove-Variable currentMobile
$textbox_Ext.Text = ""
$textbox_DID.Text = ""
$textbox_Mobile.Text = ""
}

PhoneUpdater