# Employment Tool v3
#	Tools - Add To Encryption

$adList = Get-ADUser -SearchBase "CN=Users,DC=domain,DC=com" -Filter *

	# Get username

Function AddToEncryptionForm {
$form_AddToEncryption = New-Form
$form_AddToEncryption.Size = Set-Size 350 120
$form_AddToEncryption.Text = "  Employment Tool - Add User To Encryption"
$form_AddToEncryption.SizeGripStyle = 'Hide'
$form_AddToEncryption.StartPosition = 'CenterScreen'
$form_AddToEncryption.FormBorderStyle	= 'Fixed3D'
$form_AddToEncryption.Font = "Arial,10,style=Bold"
$form_AddToEncryption.TopMost = $True
$form_AddToEncryption.KeyPreview = $True
$form_AddToEncryption.MaximizeBox = $False
$form_AddToEncryption.MinimizeBox = $False
$form_AddToEncryption.Icon = Set-Icon

$label_Username = New-Label
$label_Username.Size = Set-Size 100 20
$label_Username.Location = Set-Point 10 10
$label_Username.Text = "Username"
	$form_AddToEncryption.Controls.Add($label_Username)

$textbox_Username = New-TextBox
$textbox_Username.Size = Set-Size 200 20
$textbox_Username.Location = Set-Point 120 10
$textbox_Username.Add_TextChanged({
	$checkAD = $adList.SamAccountName |Select-String "^$($textbox_Username.Text)$" -Quiet
	If ($checkAD -eq $True) {
		$label_Username.Forecolor = "Green"
		$button_Submit.Enabled = $True
		} Else {
			$label_Username.Forecolor = "Red"
			$button_Submit.Enabled = $False
			}
	$script:theUser = $textbox_Username.Text
	})
	$form_AddToEncryption.Controls.Add($textbox_Username)
	
$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 190 50
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_AddToEncryption.Controls.Add($button_Cancel)
	$form_AddToEncryption.CancelButton = $button_Cancel
	
$button_Submit = New-Button
$button_Submit.Size = Set-Size 60 30
$button_Submit.Location = Set-Point 260 50
$button_Submit.Text = "Submit"
$button_Submit.Enabled = $False
$button_Submit.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_AddToEncryption.Controls.Add($button_Submit)
	$form_AddToEncryption.AcceptButton = $button_Submit

$fr = $form_AddToEncryption.ShowDialog()
If ($fr -eq "Cancel") { $form_AddToEncryption.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") { $form_AddToEncryption.Dispose() ; AddToEncryptionAction }
}

	# Add to group

Function AddToEncryptionAction {
$encryptionMembers = Get-ADGroupMember "Newly Encrypted Devices" | % { Get-ADUser $_ |Select SamAccountName }
If ($encryptionMembers -match $theUser) {
	Throw-Error "$($theUser) is already a member of Newly Encrypted Devices."
	} Else {
		Add-ADGroupMember "Newly Encrypted Devices" $theUser
		Write-Log "$loggedInUser added $theUser to "Newly Encrypted Devices""
		}
Run-MainMenu
}

AddToEncryptionForm