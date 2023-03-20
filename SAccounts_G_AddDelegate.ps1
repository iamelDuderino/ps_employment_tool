# Employment Tool v3.0
# 	Service Accounts - Google - Add Delegate

Function AddDelegateForm {

$form_GetDelegates = New-Form
$form_GetDelegates.Size = Set-Size 350 400
$form_GetDelegates.Text = "  Employment Tool - Google SVC Add Delegate"
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
		$label_Username.Forecolor = "Green"
		$script:svcCheck = $True
		Check-Fill
		} Else {
		$label_Username.Forecolor = "Red"
		$script:svcCheck = $False
		Check-Fill
		}
	}
	})
	$form_GetDelegates.Controls.Add($textbox_Username)

$label_DelegateTo = New-Label
$label_DelegateTo.Text = "Delegate to"
$label_DelegateTo.Size = Set-Size 100 20
$label_DelegateTo.Location = Set-Point 10 50
	$form_GetDelegates.Controls.Add($label_DelegateTo)

$richtextbox_DelegateTo = New-RichTextBox
$richtextbox_DelegateTo.Size = Set-Size 200 200
$richtextbox_DelegateTo.Location = Set-Point 120 50
$richtextbox_DelegateTo.Add_TextChanged({
If ($richtextbox_DelegateTo.Text -eq "") {
	$script:delegateCheck = $False
	$label_DelegateTo.Forecolor = "Black"
	Check-Fill
	} Else {
	If ($richtextbox_DelegateTo.Text -match ',') {
		# Start of multiple user split
		$userList2 = ($richtextbox_DelegateTo.Text).Replace("@domain.com","")
		$usersToAdd = $userList2.Split(",") -creplace "\s"
		ForEach ($user in $usersToAdd) {
			$check = $userList |Select-String "^$user\b$" -Quiet
			If (!$check){
				$label_DelegateTo.Forecolor = "Red"
				$script:delegateCheck = $False
				Check-Fill
				} Else {
					$label_DelegateTo.Forecolor = "Green"
					$script:delegateCheck = $True
					Check-Fill
					}
			}
		# End of multiple user split
		} Else {
			$user = ($richtextbox_DelegateTo.Text).Replace("@domain.com","")
			$check = $userList |Select-String "^$user\b$" -Quiet
			If ($check -eq $True) {
				$label_DelegateTo.Forecolor = "Green"
				$script:delegateCheck = $True
				Check-Fill
			} Else {
				$label_DelegateTo.Forecolor = "Red"
				$script:delegateCheck = $False
				Check-Fill
				}
			}
		}
	})
	$form_GetDelegates.Controls.Add($richtextbox_DelegateTo)
	
$label_Commas = New-Label
$label_Commas.Text = "* separate multiple users with a comma`n** can enter with or without @domain.com"
$label_Commas.Font = "Arial,8,style=italic"
$label_Commas.Forecolor = "Red"
$label_Commas.Size = Set-Size 240 40
$label_Commas.Location = Set-Point 120 260
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
	$serviceAccount = ($textbox_Username.Text).Replace("@domain.com","")
	$delegateAdditions = (($richtextbox_DelegateTo.Text).Split(",")).Replace("@domain.com","")
	$form_GetDelegates.Dispose()
	AddDelegateAction
	}
}

Function AddDelegateAction {

## Add progress bar

ForEach ($delegate in $delegateAdditions) {
	gam user $serviceAccount add delegate $delegate
	Write-Log "$loggedInUser added $delegate as a delegate to the service account: $serviceAccount"
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
If ($svcCheck -eq $True -AND $delegateCheck -eq $True) {
	$button_OK.Enabled = $True
	} Else {
		$button_OK.Enabled = $False
		}
}

$googleSVCList = (Get-ADUser -Filter * |Where {$_.DistinguishedName -notmatch "CN=Users,DC=domain,DC=com"}).SamAccountName
$userList = (Get-ADUser -Filter * -SearchBase "CN=Users,DC=domain,DC=com").SamAccountName

AddDelegateForm