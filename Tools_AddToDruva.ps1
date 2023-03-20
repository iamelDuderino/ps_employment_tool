# Employment Tool v3
#	Tools - Add To Druva

$adList = Get-ADUser -SearchBase "CN=Users,DC=domain,DC=com" -Filter *

	# Get username

Function AddToDruvaForm {
$form_AddToDruva = New-Form
$form_AddToDruva.Size = Set-Size 350 120
$form_AddToDruva.Text = "  Employment Tool - Add User To Druva"
$form_AddToDruva.SizeGripStyle = 'Hide'
$form_AddToDruva.StartPosition = 'CenterScreen'
$form_AddToDruva.FormBorderStyle	= 'Fixed3D'
$form_AddToDruva.Font = "Arial,10,style=Bold"
$form_AddToDruva.TopMost = $True
$form_AddToDruva.KeyPreview = $True
$form_AddToDruva.MaximizeBox = $False
$form_AddToDruva.MinimizeBox = $False
$form_AddToDruva.Icon = Set-Icon

$label_Username = New-Label
$label_Username.Size = Set-Size 100 20
$label_Username.Location = Set-Point 10 10
$label_Username.Text = "Username"
	$form_AddToDruva.Controls.Add($label_Username)

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
	$form_AddToDruva.Controls.Add($textbox_Username)
	
$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 190 50
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_AddToDruva.Controls.Add($button_Cancel)
	$form_AddToDruva.CancelButton = $button_Cancel
	
$button_Submit = New-Button
$button_Submit.Size = Set-Size 60 30
$button_Submit.Location = Set-Point 260 50
$button_Submit.Text = "Submit"
$button_Submit.Enabled = $False
$button_Submit.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_AddToDruva.Controls.Add($button_Submit)
	$form_AddToDruva.AcceptButton = $button_Submit

$fr = $form_AddToDruva.ShowDialog()
If ($fr -eq "Cancel") { $form_AddToDruva.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") { $form_AddToDruva.Dispose() ; AddToDruvaAction }
}

	# Add to group

Function AddToDruvaAction {
$druvaMembers = Get-ADGroupMember "Druva Cloud" | % { Get-ADUser $_ |Select SamAccountName }
If ($druvaMembers -match $theUser) {Throw-Error "$($theUser) is already a member of Druva Cloud."} Else {Add-ADGroupMember "Druva Cloud" $theUser}
Run-MainMenu
}

AddToDruvaForm