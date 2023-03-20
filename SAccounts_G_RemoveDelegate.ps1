# Employment Tool v3.0
# 	Service Accounts - Google - Remove Delegate

$googleSVCList = (Get-ADUser -Filter * -SearchBase "OU=Google SVC Accounts,DC=domain,DC=com").SamAccountName

Function RemoveDelegateForm {

$form_GetDelegates = New-Form
$form_GetDelegates.Size = Set-Size 350 400
$form_GetDelegates.Text = "  Employment Tool - Google SVC Remove Delegate"
$form_GetDelegates.SizeGripStyle = 'Hide'
$form_GetDelegates.StartPosition = 'CenterScreen'
$form_GetDelegates.FormBorderStyle = 'Fixed3D'
$form_GetDelegates.Font = "Arial,10,style=Bold"
$form_GetDelegates.TopMost = $True
$form_GetDelegates.KeyPreview = $True
$form_GetDelegates.MaximizeBox = $False
$form_GetDelegates.MinimizeBox = $False
$form_GetDelegates.Icon = Set-Icon

$label_Username = New-Label
$label_Username.Text = "Username"
$label_Username.Size = Set-Size 100 20
$label_Username.Location = Set-Point 10 20
	$form_GetDelegates.Controls.Add($label_Username)

$textbox_Username = New-TextBox
$textbox_Username.Size = Set-Size 200 20
$textbox_Username.Location = Set-Point 120 20
$textbox_Username.Add_TextChanged({
If ($textbox_Username.Text -eq "") {
	$label_Username.Forecolor = "Black"
	$script:svcCheck = $False
	Check-Fill
	} Else {
	$username = ($textbox_Username.Text).Replace("@domain.com","")
	$check = $googleSVCList |Select-String "^$username\b$" -Quiet
	If ($check -eq $True) {
		$textbox_Username.Enabled = $False
		$label_Username.Forecolor = "Green"
		$script:svcCheck = $True
		Check-Fill
		$theirEmails = GAM user $username show delegates |Select-String " Delegate Email: "
		$theirUsernames = ($theirEmails -creplace ' Delegate Email: ') -creplace '@domain.com'
		ForEach ($delegateUsername in $theirUsernames) {
			$listbox_DelegatedTo.Items.Add($delegateUsername)
			}
		$textbox_Username.Enabled = $True
		} Else {
		$listbox_DelegatedTo.Items.Clear()
		$label_Username.Forecolor = "Red"
		$script:svcCheck = $False
		Check-Fill
		}
	}
	})
	$form_GetDelegates.Controls.Add($textbox_Username)

$label_DelegatedTo = New-Label
$label_DelegatedTo.Text = "Delegated to"
$label_DelegatedTo.Size = Set-Size 100 20
$label_DelegatedTo.Location = Set-Point 10 50
	$form_GetDelegates.Controls.Add($label_DelegatedTo)

$listbox_DelegatedTo = New-ListBox
$listbox_DelegatedTo.Size = Set-Size 200 200
$listbox_DelegatedTo.Location = Set-Point 120 50
$listbox_DelegatedTo.SelectionMode = 'MultiExtended'
$listbox_DelegatedTo.Add_SelectedIndexChanged({
If ($listbox_DelegatedTo.SelectedIndex -gt -1) {
	$script:listboxCheck = $True
	Check-Fill
	} Else {
		$script:listboxCheck = $False
		Check-Fill
		}
	})
	$form_GetDelegates.Controls.Add($listbox_DelegatedTo)
	
$label_Commas = New-Label
$label_Commas.Text = "* choose multiple users with ctrl+click or shift"
$label_Commas.Font = "Arial,8,style=italic"
$label_Commas.Forecolor = "Red"
$label_Commas.Size = Set-Size 240 40
$label_Commas.Location = Set-Point 100 260
	$form_GetDelegates.Controls.Add($label_Commas)

$button_OK = New-Button
$button_OK.Text = "OK"
$button_OK.Size = Set-Size 60 40
$button_OK.Location = Set-Point 260 310
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button_OK.Enabled = $False
	$form_GetDelegates.Controls.Add($button_OK)
	$form_GetDelegates.AcceptButton = $button_OK

$button_Cancel = New-Button
$button_Cancel.Text = "Cancel"
$button_Cancel.Size = Set-Size 60 40
$button_Cancel.Location = Set-Point 190 310
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_GetDelegates.Controls.Add($button_Cancel)
	$form_GetDelegates.CancelButton = $button_Cancel

$fr = $form_GetDelegates.ShowDialog()

If ($fr -eq "Cancel") {
	$form_GetDelegates.Dispose()
	Run-MainMenu
	}
	
If ($fr -eq "OK") { 
	$serviceAccount = $textbox_Username.Text
	$delegatesToRemove = $listbox_DelegatedTo.SelectedItems
	$form_GetDelegates.Dispose()
	RemoveDelegateAction
	}
}

Function RemoveDelegateAction {
ForEach ($delegateToRemove in $delegatesToRemove) {
	GAM user $serviceAccount delete delegate $delegateToRemove
	Write-Log "$loggedInUser removed the delegate $delegateToRemove from the service account: $serviceAccount"
	}
$theirEmails = GAM user $serviceAccount show delegates |Select-String " Delegate Email: "
$theirUsernames = ($theirEmails -creplace ' Delegate Email: ') -creplace '@domain.com'
$descr = "Delegated to: "
ForEach ($username in $theirUsernames) {
	$descr += $username
	$descr += ', '
	}
If ($descr -ne "$descr,..$") {
	$descr = $descr -Replace ",.$"
	}
Set-ADUser $serviceAccount -Description "$descr"
Run-MainMenu
}

Function Check-Fill {
If ($svcCheck -eq $True -AND $listboxCheck -eq $True) {
	$button_OK.Enabled = $True
	} Else {
		$button_OK.Enabled = $False
		}
}

RemoveDelegateForm