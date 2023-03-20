# Employment Tool v3.0
# 	Tools - Create JIRA & JIRA-Test Account

$adList = Get-ADUser -Filter * -SearchBase "CN=Users,DC=domain,DC=com"

	# Gather information for user to create WebEx account for

Function Start-Form {
$form_Query = New-Form
$form_Query.Size = Set-Size 350 120
$form_Query.Text = "  Employment Tool - Create JIRA/Test Account"
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
	})
	$form_Query.Controls.Add($textbox_Username)

$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 190 50
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_Query.Controls.Add($button_Cancel)
	$form_Query.CancelButton = $button_Cancel
	
$button_Submit = New-Button
$button_Submit.Size = Set-Size 60 30
$button_Submit.Location = Set-Point 260 50
$button_Submit.Text = "Submit"
$button_Submit.Enabled = $False
$button_Submit.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_Query.Controls.Add($button_Submit)
	$form_Query.AcceptButton = $button_Submit

$fr = $form_Query.ShowDialog()
If ($fr -eq "Cancel") { $form_Query.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") {
	$script:theUser = $textbox_Username.Text
	$userProps = Get-ADUser $theUser -Properties *
	$script:theDisplayName = $userProps.GivenName + " " + $userProps.Surname
	$script:theEmail = $userProps.Mail
	$form_Query.Dispose()	
	CreateJIRAAccount
	}
}

	# If there are no sync errors, create the JIRA and JIRA-Test accounts

Function CreateJIRAAccount {
	# JIRA & JIRA-Test API

$JIRAURI = "https://jira.domain.com/rest/api/2/user/search?username=$theUser"
$JIRATestURI = "https://jira-test.domain.com/rest/api/2/user/search?username=$theUser"
$AuthKey = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$AuthPassword = Get-Content "C:\Scripts\Text Files\JIRA-SVCITAPIAdminCredentials.txt" |ConvertTo-SecureString -Key $AuthKey
$AutoCreds = New-Object System.Management.Automation.PSCredential -ArgumentList "svcitapiadmin",$AuthPassword
$Creds = $AutoCreds.GetNetworkCredential()
$Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $Creds.Username,$Creds.Password)))
$checkJIRA = Invoke-RestMethod -URI $JIRAURI -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)}
$checkJIRATest = Invoke-RestMethod -URI $JIRATestURI -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)}

	# JIRA
	
If (!$checkJIRA) {
	$ProvisionJIRA = @{
		"name" = "$theUser"
		"password" = "Day1@domain!"
		"displayName" = "$theDisplayName"
		"emailAddress" = "$theEmail"
	}
	$JSON = $ProvisionJIRA |ConvertTo-JSON
	$CreateURI = "https://jira.domain.com/rest/api/2/user"
	Invoke-RestMethod -URI $CreateURI -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} -ContentType "application/json" -Body $JSON -Method POST
	$JIRAGroups = "engineering" , "fisheye"
	$userToAdd = @{
		"name" = "$theUser"
	}
	$JSON = $userToAdd |ConvertTo-JSON
	ForEach ($JIRAGroup in $JIRAGroups) {
		$JIRAGroupURI = "https://jira.domain.com/rest/api/2/group/user?groupname=$JIRAGroup"
		Invoke-RestMethod -URI $JIRAGroupURI -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} -ContentType "application/json" -Body $JSON -Method POST
	}
	Throw-Error "User account setup. Default password is:`n`nDay1@domain!"
} Else { Throw-Error "Existing user with username $($theUser) was found in JIRA system." }
	
	# JIRA-Test
	
If (!$checkJIRATest) {
	$ProvisionJIRA = @{
		"name" = "$theUser"
		"displayName" = "$theDisplayName"
		"emailAddress" = "$theEmail"
	}
	$JSON = $ProvisionJIRA |ConvertTo-JSON
	$CreateURI = "https://jira-test.domain.com/rest/api/2/user"
	Invoke-RestMethod -URI $CreateURI -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} -ContentType "application/json" -Body $JSON -Method POST
	$JIRAGroups = "engineering" , "fisheye"
	$userToAdd = @{
		"name" = "$theUser"
	}
	$JSON = $userToAdd |ConvertTo-JSON
	ForEach ($JIRAGroup in $JIRAGroups) {
		$JIRAGroupURI = "https://jira-test.domain.com/rest/api/2/group/user?groupname=$JIRAGroup"
		Invoke-RestMethod -URI $JIRAGroupURI -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} -Method POST -ContentType "application/json" -Body $JSON
		}
	} Else { Throw-Error "Existing user with username $($theUser) was found in JIRA-Test system." }
Run-MainMenu
}

Start-Form