# Employment Tool v3.0
# 	Service Accounts - Google - Creation

	## Gather information for the service account creation

Function GoogleSVC {

$ADList = Get-ADUser -Filter *

$form_GoogleSVC = New-Form
$form_GoogleSVC.Size = Set-Size 420 270
$form_GoogleSVC.Text = "  Employment Tool - Google SVC Creation"
$form_GoogleSVC.SizeGripStyle = 'Hide'
$form_GoogleSVC.StartPosition = 'CenterScreen'
$form_GoogleSVC.FormBorderStyle = 'Fixed3D'
$form_GoogleSVC.Font = "Arial,10,style=Bold"
$form_GoogleSVC.TopMost = $True
$form_GoogleSVC.KeyPreview = $True
$form_GoogleSVC.MaximizeBox = $False
$form_GoogleSVC.MinimizeBox = $False
$form_GoogleSVC.Icon = Set-Icon

$label_FirstName = New-Label
$label_FirstName.Size = Set-Size 100 20
$label_FirstName.Location = Set-Point 10 10
$label_FirstName.Text = "First Name"
	$form_GoogleSVC.Controls.Add($label_FirstName)
	
$textbox_FirstName = New-Textbox
$textbox_FirstName.Size = Set-Size 250 20
$textbox_FirstName.Location = Set-Point 150 10
$textbox_FirstName.Add_TextChanged({Validate-Form})
	$form_GoogleSVC.Controls.Add($textbox_FirstName)

$label_LastName = New-Label
$label_LastName.Size = Set-Size 100 20
$label_LastName.Location = Set-Point 10 40
$label_LastName.Text = "Last Name"
	$form_GoogleSVC.Controls.Add($label_LastName)

$textbox_LastName = New-Textbox
$textbox_LastName.Size = Set-Size 250 20
$textbox_LastName.Location = Set-Point 150 40
$textbox_LastName.Add_TextChanged({Validate-Form})
	$form_GoogleSVC.Controls.Add($textbox_LastName)

$label_Username = New-Label
$label_Username.Size = Set-Size 100 20
$label_Username.Location = Set-Point 10 70
$label_Username.Text = "Username"
	$form_GoogleSVC.Controls.Add($label_Username)

$textbox_Username = New-Textbox
$textbox_Username.Size = Set-Size 250 20
$textbox_Username.Location = Set-Point 150 70
$textbox_Username.MaxLength = 20
$textbox_Username.Add_TextChanged({
	If ($textbox_Username.Text -match '@') {
		Throw-Error "Please enter the username without @sevone.com"
		$textbox_Username.Text = ($textbox_Username.Text).Trim('@')
		$textbox_Username.SelectionStart = $textbox_Username.TextLength
		}
	$checkAD = $ADList.SamAccountName |Select-String "^$($textbox_Username.Text)\b$" -Quiet
	If ($checkAD -eq $True) {
		$label_Username.Forecolor = "Red"
		$script:usernameValid = $False
		Validate-Form
		} Else {
			$label_Username.Forecolor = "Green"
			$script:usernameValid = $True
			Validate-Form
			}
	If ($textbox_Username.Text -eq "") {$textbox_Username.Forecolor = "Black"}
	})
	$form_GoogleSVC.Controls.Add($textbox_Username)

$label_Delegates = New-Label
$label_Delegates.Size = Set-Size 100 20
$label_Delegates.Location = Set-Point 10 100
$label_Delegates.Text = "Delegate To"
	$form_GoogleSVC.Controls.Add($label_Delegates)

$textbox_Delegates = New-RichTextbox
$textbox_Delegates.Size = Set-Size 250 60
$textbox_Delegates.Location = Set-Point 150 100
$textbox_Delegates.Add_TextChanged({Validate-Form})
	$form_GoogleSVC.Controls.Add($textbox_Delegates)
	
$label_Comma = New-Label
$label_Comma.Size = Set-Size 260 20
$label_Comma.Location = Set-Point 160 170
$label_Comma.Forecolor = "Red"
$label_Comma.Font = "Arial,8,style=Italic"
$label_Comma.Text = "* Separate multiple delegates with a comma"
	$form_GoogleSVC.Controls.Add($label_Comma)

$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 270 200
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_GoogleSVC.Controls.Add($button_Cancel)
	$form_GoogleSVC.CancelButton = $button_Cancel

$button_OK = New-Button
$button_OK.Size = Set-Size 60 30
$button_OK.Location = Set-Point 340 200
$button_OK.Text = "OK"
$button_OK.Enabled = $False
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button_OK.Add_Click({
	$script:theFirstName = $textbox_FirstName.Text
	$script:theLastName = $textbox_LastName.Text
	$script:theDisplayName = "$theFirstName $theLastName"
	$script:theUser = $textbox_Username.Text
	$script:theEmail = "$theUser@sevone.com"
	$script:theDelegates = ($textbox_Delegates.Text).Split(",") -creplace '\s' | % { $_ -creplace '@sevone.com' }
	})
	$form_GoogleSVC.Controls.Add($button_OK)
	$form_GoogleSVC.AcceptButton = $button_OK

$fr = $form_GoogleSVC.ShowDialog()
If ($fr -eq "Cancel") { $form_GoogleSVC.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") { $form_GoogleSVC.Dispose() ; Create-GoogleSVC }
}

	# Confirm that form is filled out with required information

Function Validate-Form {
If ($textbox_FirstName.Text -AND $textbox_LastName.Text -AND $textbox_Delegates.Text -AND $usernameValid -eq $True) {
	$button_OK.Enabled = $True
	} Else {
		$button_OK.Enabled = $False
		}
}

	# Create the Google Service Account

Function Create-GoogleSVC {
$form_GSync = New-Form
$form_GSync.Size = Set-Size 200 100
$form_GSync.Text = "  Employment Tool"
$form_GSync.SizeGripStyle = 'Hide'
$form_GSync.StartPosition = 'CenterScreen'
$form_GSync.FormBorderStyle = 'Fixed3D'
$form_GSync.Font = "Arial,10,style=Bold"
$form_GSync.TopMost = $True
$form_GSync.KeyPreview = $True
$form_GSync.MaximizeBox = $False
$form_GSync.MinimizeBox = $False
$form_GSync.Icon = Set-Icon

$label_GSync = New-Label
$label_GSync.Size = Set-Size 150 20
$label_GSync.Location = Set-Point 20 20
$label_GSync.Text = "Syncing to AD"
	$form_GSync.Controls.Add($label_GSync)

$form_GSync.Show()

	## Active Directory Account Creation

$theGeneratedPW = New-RandomPassword -Length 20 -Uppercase -Lowercase -Symbols -Numbers
$thePassword = ConvertTo-SecureString $theGeneratedPW -AsPlainText -Force
$theOU = "OU=Google SVC Accounts,DC=sevone,DC=com"
$theGoogleOU = "/003 - Google SVC Accounts"
$theUserProperties = @{
	'SamAccountName' = "$theUser"
	'UserPrincipalName' = "$theEmail"
	'Name' = "$theDisplayName"
	'GivenName' = "$theFirstName"
	'Surname' = "$theLastName"
	'EmailAddress' = "$theEmail"
	'DisplayName' = "$theDisplayName"
	'AccountPassword' = $thePassword
	'ChangePasswordAtLogon' = $False
	'Enabled' = $False
	'Path' = "$theOU"
	'PasswordNeverExpires' = $True
	}
New-ADUser @theUserProperties
Sleep 2
	
	## Google Account Creation

Sync-Google
Sleep 1
Do {
	$form_GSync.Refresh()
	$label_GSync.Text = "Syncing to google."
	Sleep 1
	$form_GSync.Refresh()
	$checkProcess = Get-Process |Select ProcessName |Where {$_.ProcessName -match "sync-cmd"}
	$label_GSync.Text = "Syncing to google.."
	Sleep 1
	$form_GSync.Refresh()
	$checkProcess = Get-Process |Select ProcessName |Where {$_.ProcessName -match "sync-cmd"}
	$label_GSync.Text = "Syncing to google..."
	Sleep 1
	$form_GSync.Refresh()
	$checkProcess = Get-Process |Select ProcessName |Where {$_.ProcessName -match "sync-cmd"}
	$label_GSync.Text = "Syncing to google...."
	Sleep 1
	$form_GSync.Refresh()
	$checkProcess = Get-Process |Select ProcessName |Where {$_.ProcessName -match "sync-cmd"}
	} Until (!$checkProcess)

$label_GSync.Text = "Done!"
$form_GSync.Refresh()
Sleep 2

Write-Log "$loggedInUser created the google service account $theDisplayName ($theEmail)"

	## Delegate & Update

GAM Update User $theUser Org $theGoogleOU
GAM Update User $theUser Password $thePassword
GAM User $theUser Profile Unshared

$label_GSync.Text = "Google updated!"
$form_GSync.Refresh()

$mailBody = "
Hello,

You have been added as a delegate to the account $theDisplayName ($theEmail).

Please refresh your gmail and this account can be accessed by clicking your avatar circle in the top right.

Regards,

-IT Team
"

ForEach ($theDelegate in $theDelegates) {
	GAM user $theUser delegate to $theDelegate
	Write-Log "$loggedInUser delegated $theUser to $theDelegate"
	Send-MailMessage -From "SevOne Admin <sevoneadmin@sevone.com>" -To "$theDelegate@sevone.com" -Subject "[SVC Account] You've Been Added As A Delegate!" -Body $mailBody -SMTPServer 'aspmx.l.google.com' -Port 25
	}

$label_GSync.Text = "Account delegated!"
$form_GSync.Refresh()

$theDelegatesList = gam user $theUser show delegates |Select-String " Delegate Email" | % { $_ -creplace '.*\s' } | % { $_ -creplace '@sevone.com' }
$n = 0
If ($theDelegatesList.Count -eq '1') {
	$theDescr = "Delegated to: $theDelegatesList"
	Set-ADUser $theUser -Description $theDescr
	}
If ($theDelegatesList.Count -gt '1') {
	$num = $theDelegatesList.Count
	$theDescr = "Delegated to: "
Do {
	If ($n -lt $num -AND $n -ne $num-1) {
		$theDescr += "$($theDelegatesList[$n]), "
		}
	If ($n -eq $num-1) {
		$theDescr += "$($theDelegatesList[$n])"
		}
	$n++
} Until ($n -eq $num)
	Set-ADUser $theUser -Description $theDescr
	}
$label_GSync.Text = "Done!"
$form_GSync.Refresh()
$form_GSync.Dispose()
Run-MainMenu
}

GoogleSVC