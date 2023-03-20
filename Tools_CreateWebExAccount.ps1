# Employment Tool v3.0
# 	Tools - Create WebEx Account

$adList = Get-ADUser -Filter * -SearchBase "CN=Users,DC=sevone,DC=com"

	# Gather information for user to create WebEx account for
	
Function Start-Form {
$form_Query = New-Form
$form_Query.Size = Set-Size 340 170
$form_Query.Text = "  Employment Tool - Create WebEx Account"
$form_Query.SizeGripStyle = 'Hide'
$form_Query.StartPosition = 'CenterScreen'
$form_Query.FormBorderStyle = 'Fixed3D'
$form_Query.Font = "Arial,10,style=Bold"
$form_Query.TopMost = $True
$form_Query.KeyPreview = $True
$form_Query.MaximizeBox = $False
$form_Query.MinimizeBox = $False
$form_Query.Icon = Set-Icon

$label_Username = New-Label
$label_Username.Size = Set-Size 100 20
$label_Username.Location = Set-Point 10 10
$label_Username.Text = "Username"
	$form_Query.Controls.Add($label_Username)
	
$textbox_Username = New-Textbox
$textbox_Username.Size = Set-Size 200 20
$textbox_Username.Location = Set-Point 120 10
$textbox_Username.Add_TextChanged({
	$checkAD = $adList.SamAccountName |Select-String "^$($textbox_Username.Text)$" -Quiet
	If ($checkAD -eq $True){
		$label_Username.Forecolor = "Green"
		$button_Submit.Enabled = $True
		} Else {
			$label_Username.Forecolor = "Red"
			$button_Submit.Enabled = $False
			}
	If ($textbox_Username.Text -eq "") {
		$label_Username.Forecolor = "Black"
		}
	$script:theUser = $textbox_Username.Text
	})
	$form_Query.Controls.Add($textbox_Username)
	
$checkbox_MeetingCenter = New-Checkbox
$checkbox_MeetingCenter.Size = Set-Size 150 20
$checkbox_MeetingCenter.Location = Set-Point 10 40
$checkbox_MeetingCenter.Text = "Meeting Center"
$checkbox_MeetingCenter.Enabled = $False
$checkbox_MeetingCenter.Checked = $True
	$form_Query.Controls.Add($checkbox_MeetingCenter)

$checkbox_EventCenter = New-Checkbox
$checkbox_EventCenter.Size = Set-Size 150 20
$checkbox_EventCenter.Location = Set-Point 170 40
$checkbox_EventCenter.Text = "Event Center"
	$form_Query.Controls.Add($checkbox_EventCenter)

$checkbox_SupportCenter = New-Checkbox
$checkbox_SupportCenter.Size = Set-Size 150 20
$checkbox_SupportCenter.Location = Set-Point 10 70
$checkbox_SupportCenter.Text = "Support Center"
	$form_Query.Controls.Add($checkbox_SupportCenter)

$checkbox_TrainingCenter = New-Checkbox
$checkbox_TrainingCenter.Size = Set-Size 150 20
$checkbox_TrainingCenter.Location = Set-Point 170 70
$checkbox_TrainingCenter.Text = "Training Center"
	$form_Query.Controls.Add($checkbox_TrainingCenter)
	
$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 190 100
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_Query.Controls.Add($button_Cancel)
	$form_Query.CancelButton = $button_Cancel
	
$button_Submit = New-Button
$button_Submit.Size = Set-Size 60 30
$button_Submit.Location = Set-Point 260 100
$button_Submit.Text = "Submit"
$button_Submit.Enabled = $False
$button_Submit.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_Query.Controls.Add($button_Submit)
	$form_Query.AcceptButton = $button_Submit

$fr = $form_Query.ShowDialog()
If ($fr -eq "Cancel") { $form_Query.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") {
	$script:thePerms = ""
	If ($checkbox_EventCenter.Checked -eq $True) {
		If ($thePerms -eq "") {
		$thePerms += "<meetingType>220</meetingType>"
		} Else {
			$thePerms += "
				<meetingType>220</meetingType>"
			}
	}
	If ($checkbox_SupportCenter.Checked -eq $True) {
		If ($thePerms -eq "") {
		$thePerms += "<meetingType>565</meetingType>"
		} Else {
			$thePerms += "
				<meetingType>565</meetingType>"
			}
	}
	If ($checkbox_TrainingCenter.Checked -eq $True) {
		If ($thePerms -eq "") {
		$thePerms += "<meetingType>522</meetingType>"
		} Else {
			$thePerms += "
				<meetingType>522</meetingType>"
			}
	}
	$form_Query.Dispose()
	CreateWebExAccount
	}
}

	# If there are no sync errors, create the WebEx account
	
Function CreateWebExAccount {
$userInfo = Get-ADUser $theUser -Properties *
$theFirstName = $userInfo.GivenName
$theLastName = $userInfo.Surname
$theEmail = "$theUser@sevone.com"
$URI = 'https://sevone.webex.com/WBXService/XMLService'
$ContentType = 'text/xml'
$Method = 'POST'

$CheckExistingBody = (Get-Content "C:\Scripts\Text Files\GetWebExUser.xml").Replace("THE_USERNAME_TO_CHECK","$theUser")

$checkExistingResults = (Invoke-WebRequest -URI $URI -Method $Method -ContentType $ContentType -Body $CheckExistingBody).Content

If ($checkExistingResults -match "FAILURE") { # Proceed to provision because there is no user by that name
	$BodyFixFirst = (Get-Content "C:\Scripts\Text Files\CreateWebExUser.xml").Replace("THE_FIRSTNAME","$theFirstName")
	$BodyFixLast = $BodyFixFirst.Replace("THE_LASTNAME","$theLastName")
	$BodyFixUsername = $BodyFixLast.Replace("THE_USERNAME","$theUser")
	$BodyFixEmail = $BodyFixUsername.Replace("THE_EMAIL","$theEmail")
	If ($thePerms) {
		$BodyFixPermissions = $BodyFixEmail.Replace("THE_MEETING_PERMS","$thePerms")
		} Else {
			$BodyFixPermissions = $BodyFixEmail |Where {$_ -notmatch "THE_MEETING_PERMS"}
			}
	$ProvisionBody = $BodyFixPermissions
	Invoke-WebRequest -URI $URI -Method $Method -ContentType $ContentType -Body $ProvisionBody
	Throw-Error "$($theUser) was provisioned in WebEx."
	} Else { # Throw-Error advising username already used
		Throw-Error "The username $($theUser) already exists in the WebEx System!"
		}
Run-MainMenu
}

Start-Form