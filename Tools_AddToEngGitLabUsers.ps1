# Employment Tool v3
#	Tools - Add To EngGitLabUsers

$adList = Get-ADUser -SearchBase "CN=Users,DC=domain,DC=com" -Filter *

	# Get username

Function AddToSRCForm {
$form_AddToSRC = New-Form
$form_AddToSRC.Size = Set-Size 350 120
$form_AddToSRC.Text = "  Employment Tool - Add User To GitLab (SRC)"
$form_AddToSRC.SizeGripStyle = 'Hide'
$form_AddToSRC.StartPosition = 'CenterScreen'
$form_AddToSRC.FormBorderStyle	= 'Fixed3D'
$form_AddToSRC.Font = "Arial,10,style=Bold"
$form_AddToSRC.TopMost = $True
$form_AddToSRC.KeyPreview = $True
$form_AddToSRC.MaximizeBox = $False
$form_AddToSRC.MinimizeBox = $False
$form_AddToSRC.Icon = Set-Icon

$label_Username = New-Label
$label_Username.Size = Set-Size 100 20
$label_Username.Location = Set-Point 10 10
$label_Username.Text = "Username"
	$form_AddToSRC.Controls.Add($label_Username)

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
	$form_AddToSRC.Controls.Add($textbox_Username)
	
$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 190 50
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_AddToSRC.Controls.Add($button_Cancel)
	$form_AddToSRC.CancelButton = $button_Cancel
	
$button_Submit = New-Button
$button_Submit.Size = Set-Size 60 30
$button_Submit.Location = Set-Point 260 50
$button_Submit.Text = "Submit"
$button_Submit.Enabled = $False
$button_Submit.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_AddToSRC.Controls.Add($button_Submit)
	$form_AddToSRC.AcceptButton = $button_Submit

$fr = $form_AddToSRC.ShowDialog()
If ($fr -eq "Cancel") { $form_AddToSRC.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") { $form_AddToSRC.Dispose() ; AddToSRCAction }
}

	# Add to group

Function AddToSRCAction {
$srcMembers = Get-ADGroupMember "EngGitLabUsers" | % { Get-ADUser $_ |Select SamAccountName }
If ($srcMembers -match $theUser) {
	Throw-Error "$($theUser) is already a member of EngGitLabUsers."
	} Else {
		Add-ADGroupMember "EngGitLabUsers" $theUser
		Write-Log "$loggedInUser added $theUser to "EngGitLabUsers""
		}
Run-MainMenu
}

AddToSRCForm