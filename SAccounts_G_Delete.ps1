# Employment Tool v3.0
# 	Service Accounts - Google - Deletion

	## Gather information for the service account deletion

Function GoogleSVC {

$ADList = Get-ADUser -Filter * -SearchBase "OU=Google SVC Accounts,DC=domain,DC=com"

$form_GoogleSVC = New-Form
$form_GoogleSVC.Size = Set-Size 380 120
$form_GoogleSVC.Text = "  Employment Tool - Google SVC Deletion"
$form_GoogleSVC.SizeGripStyle = 'Hide'
$form_GoogleSVC.StartPosition = 'CenterScreen'
$form_GoogleSVC.FormBorderStyle = 'Fixed3D'
$form_GoogleSVC.Font = "Arial,10,style=Bold"
$form_GoogleSVC.TopMost = $True
$form_GoogleSVC.KeyPreview = $True
$form_GoogleSVC.MaximizeBox = $False
$form_GoogleSVC.MinimizeBox = $False
$form_GoogleSVC.Icon = Set-Icon

$label_Username = New-Label
$label_Username.Size = Set-Size 150 20
$label_Username.Location = Set-Point 10 10
$label_Username.Text = "Username"
	$form_GoogleSVC.Controls.Add($label_Username)

$textbox_Username = New-Textbox
$textbox_Username.Size = Set-Size 200 20
$textbox_Username.Location = Set-Point 160 10
$textbox_Username.Add_TextChanged({
	$script:theUser = $textbox_Username.Text
	If ($textbox_Username.Text -match '@') {
		Throw-Error "Please enter the username without @domain.com"
		$textbox_Username.Text = ($textbox_Username.Text).Trim('@')
		$textbox_Username.SelectionStart = $textbox_Username.TextLength
		}
	$checkAD = $ADList.SamAccountName |Select-String "^$($textbox_Username.Text)\b$" -Quiet
	If ($checkAD -eq $True) {
		$label_Username.Forecolor = "Green"
		$script:usernameValid = $True
		Validate-Form
		} Else {
			$label_Username.Forecolor = "Red"
			$script:usernameValid = $False
			Validate-Form
			}
	If ($textbox_Username.Text -eq " ") {$textbox_Username.Forecolor = "Black"}
	})
	$form_GoogleSVC.Controls.Add($textbox_Username)

$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 220 50
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_GoogleSVC.Controls.Add($button_Cancel)
	$form_GoogleSVC.CancelButton = $button_Cancel

$button_OK = New-Button
$button_OK.Size = Set-Size 60 30
$button_OK.Location = Set-Point 300 50
$button_OK.Text = "OK"
$button_OK.Enabled = $False
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_GoogleSVC.Controls.Add($button_OK)
	$form_GoogleSVC.AcceptButton = $button_OK

$fr = $form_GoogleSVC.ShowDialog()
If ($fr -eq "Cancel") { $form_GoogleSVC.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") { $form_GoogleSVC.Dispose() ; Delete-GoogleSVC }
}

	# Confirm that the form is filled out with the required information

Function Validate-Form {
If ($usernameValid -eq $True) {
	$button_OK.Enabled = $True
	} Else {
		$button_OK.Enabled = $False
		}
}

	# Delete Google Service account

Function Delete-GoogleSVC {
$theUserProps = Get-ADUser $theUser -Properties *
$r = Throw-Error "You are about to delete the account $($theUserProps.GivenName) $($theUserProps.Surname) $($theUserProps.Mail).`n`nPress OK to continue.`n`nPress Cancel to abort."
If ($r -eq "Cancel") { GoogleSVC }
If ($r -eq "OK") {
	Remove-ADUser $theUser
	Write-Log "$loggedInUser deleted the Google SVC Account for $theUser"
	Run-MainMenu
	}
}

GoogleSVC