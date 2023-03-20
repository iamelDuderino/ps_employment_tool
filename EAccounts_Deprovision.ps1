# Employment Tool v3
#	Employee Accounts - De-Provision

	# Gather information for de-provisioning

Function Start-Form {

$ADList = Get-ADUser -Filter * -SearchBase "CN=Users,DC=sevone,DC=com" |Select SamAccountName

$form_DeProv = New-Form
$form_DeProv.Size = Set-Size 480 200
$form_DeProv.Text = "  Employment Tool - Deprovision"
$form_DeProv.SizeGripStyle = 'Hide'
$form_DeProv.StartPosition = 'CenterScreen'
$form_DeProv.FormBorderStyle = 'Fixed3D'
$form_DeProv.Font = "Arial,10,style=Bold"
$form_DeProv.TopMost = $True
$form_DeProv.KeyPreview	= $True
$form_DeProv.MaximizeBox = $False
$form_DeProv.MinimizeBox = $False
$form_DeProv.Icon = Set-Icon

$label_Username = New-Label
$label_Username.Size = Set-Size 130 20
$label_Username.Location = Set-Point 10 10
$label_Username.Text = "Username"
	$form_DeProv.Controls.Add($label_Username)

$textbox_Username = New-Textbox
$textbox_Username.Size = Set-Size 300 20
$textbox_Username.Location = Set-Point 150 10
$textbox_Username.Add_TextChanged({
	$checkAD = $ADList.SamAccountName |Select-String "^$($textbox_Username.Text)\b$" -Quiet
	If ($checkAD -eq $True) {
		$label_Username.ForeColor = "Green"
		$script:usernameCorrect = $True
		Validate-Form
		}
	Else {
		$label_Username.ForeColor = "Red"
		$script:usernameCorrect = $False
		Validate-Form
		}
	})
	$form_DeProv.Controls.Add($textbox_Username)

$label_Manager = New-Label
$label_Manager.Size = Set-Size 130 20
$label_Manager.Location = Set-Point 10 40
$label_Manager.Text = "Manager"
	$form_DeProv.Controls.Add($label_Manager)

$textbox_Manager = New-Textbox
$textbox_Manager.Size = Set-Size 300 20
$textbox_Manager.Location = Set-Point 150 40
$textbox_Manager.Add_TextChanged({
	$checkAD = $ADList.SamAccountName |Select-String "^$($textbox_Manager.Text)\b$" -Quiet
	If ($checkAD -eq $True) {
		$label_Manager.ForeColor = "Green"
		$script:managerCorrect = $True
		Validate-Form
		}
	Else {
		$label_Manager.ForeColor = "Red"
		$script:managerCorrect = $False
		Validate-Form
		}
	})
	$form_DeProv.Controls.Add($textbox_Manager)

$checkbox_ShouldRetain = New-Checkbox
$checkbox_ShouldRetain.Size = Set-Size 480 20
$checkbox_ShouldRetain.Location = Set-Point 10 70
$checkbox_ShouldRetain.Text = "Should the account be retained for more than 90 days?"
	$form_DeProv.Controls.Add($checkbox_ShouldRetain)

$checkbox_ShouldSetOOO = New-Checkbox
$checkbox_ShouldSetOOO.Size = Set-Size 480 20
$checkbox_ShouldSetOOO.Location = Set-Point 10 100
$checkbox_ShouldSetOOO.Text = "Should the OOO response be set?"
	$form_DeProv.Controls.Add($checkbox_ShouldSetOOO)


$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 70 30
$button_Cancel.Location = Set-Point 300 130
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_DeProv.Controls.Add($button_Cancel)
	$form_DeProv.CancelButton = $button_Cancel

$button_Submit = New-Button
$button_Submit.Size = Set-Size 70 30
$button_Submit.Location = Set-Point 380 130
$button_Submit.Text = "OK"
$button_Submit.Enabled = $False
$button_Submit.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_DeProv.Controls.Add($button_Submit)
	$form_DeProv.AcceptButton = $button_Submit

$fr = $form_DeProv.ShowDialog()
If ($fr -eq "Cancel") { Run-MainMenu }
If ($fr -eq "OK") {
	$script:theUser = $textbox_Username.Text
	$script:theManager = $textbox_Manager.Text
	If ($checkbox_ShouldSetOOO.CheckState -eq "Checked"){$script:shouldSetOOO = "Yes"} Else {$script:shouldSetOOO = "No"}
	If ($checkbox_ShouldRetain.CheckState -eq "Checked"){$script:shouldRetain = "Yes"} Else {$script:shouldRetain = "No"}
	Confirm-Deprov
	}
}

	# Enable the submit button only once all fields have values

Function Validate-Form {
	If (
		$usernameCorrect -eq $True -AND
		$managerCorrect -eq $True
		) {$button_Submit.Enabled = $True} Else {$button_Submit.Enabled = $False}
}

	# Prompt for confirmation

Function Confirm-Deprov {
$theUserInformation = Get-ADUser $theUser
$theManagerInformation = Get-ADUser $theManager
$theUserInfo = "$($theUserInformation.GivenName) $($theUserInformation.Surname) ($($theUserInformation.UserPrincipalName))"
$theManagerInfo = "$($theManagerInformation.GivenName) $($theManagerInformation.Surname) ($($theManagerInformation.UserPrincipalName))"
$theManagerInformation = Get-ADUser $theManager
$confirm = [System.Windows.Forms.MessageBox]::SHOW(
"Employee: $($theUserInfo)`n
Delegated To: $($theManagerInfo)`n
Set OOO: $($shouldSetOOO)`n
Should Retain: $($shouldRetain)`n`n
Are you sure you wish to deprovision?",
	"Warning",
	"OKCancel",
	"Warning"
)
If ($confirm -eq "OK") { Deprovision-User }
If ($confirm -eq "Cancel") { Start-Form }
}

	# Begin deprovisioning process

Function Deprovision-User {
$form_Progress = New-Form
$form_Progress.Size = Set-Size 290 140
$form_Progress.Text = "  Employment Tool"
$form_Progress.SizeGripStyle = 'Hide'
$form_Progress.StartPosition = 'CenterScreen'
$form_Progress.FormBorderStyle = 'Fixed3D'
$form_Progress.Font = "Arial,10,style=Bold"
$form_Progress.TopMost = $True
$form_Progress.KeyPreview = $True
$form_Progress.MaximizeBox = $False
$form_Progress.MinimizeBox = $False
$form_Progress.Icon = Set-Icon

$label_Progress = New-Label
$label_Progress.Size = Set-Size 250 20
$label_Progress.Location = Set-Point 10 10
$label_Progress.Text = "Beginning to deprovision..."
	$form_Progress.Controls.Add($label_Progress)

$progressbar_Progress = New-ProgressBar
$progressbar_Progress.Size = Set-Size 250 20
$progressbar_Progress.Location = Set-Point 10 40
$progressbar_Progress.Value = 0
	$form_Progress.Controls.Add($progressbar_Progress)

$form_Progress.Show()

	# Remove AD Groups

Add-ADGroupMember -Identity "Former Employees" -Members $theUser
$newPrimaryGroup = Get-ADGroup "Former Employees" -Properties @("primaryGroupToken")
Get-ADUser $theUser |Set-ADUser -Replace @{primaryGroupID=$newPrimaryGroup.primaryGroupToken}
Get-ADPrincipalGroupMembership -Identity $theUser |select "name" |?{$_ -notmatch "Former Employees"}|%{Remove-ADPrincipalGroupMembership -Identity $theUser -MemberOf $_.Name}
	$label_Progress.Text = "AD Groups Removed..."
	$progressbar_Progress.Value = 10
	$form_Progress.Refresh()

	# Reset passwords and disable accounts

$newPW = New-RandomPassword -Length 15 -Symbols -Lowercase -Uppercase -Numbers
$securePW = ConvertTo-SecureString $newPW -AsPlainText -Force
Set-ADAccountPassword -Identity $theUser -NewPassword $securePW -PassThru
Set-ADUser $theUser -Enabled $False
GAM update user $theUser password random
GAM user $theUser deprovision
GAM user $theUser profile unshared
GAM user $theUser delegate to $theManager
Set-ADUser $theUser -Description "Delegated to: $theManager"
GAM calendar $theUser add editor $theManager
GAM create datatransfer $theUser drive $theManager
gam print mobile query $theUser | %{ $_ -creplace ',.*' } |Select -Skip 1 |% { gam delete mobile "$_" }
	$label_Progress.Text = "Accounts disabled..."
	$progressbar_Progress.Value = 30
	$form_Progress.Refresh()
	
	# Remove All Google Groups

GAM info user $theUser |Findstr "@sevone.com>" | % { $_ -creplace '^.*<' } | % { $_.Replace(">","") } | % { gam update group $_ remove user $theUser }
	$label_Progress.Text = "Google Groups Removed..."
	$progressbar_Progress.Value = 40
	$form_Progress.Refresh()

	# Set Out of Office

If ($shouldSetOOO -eq "Yes") {
	$managerInfo = Get-ADUser $theManager -Properties *
	GAM user $theUser Vacation On Subject "[Out of Office]" Message "Hello,`n`nThank you for contacting me. Please reach out to $($managerInfo.GivenName) $($managerInfo.Surname) ($($managerInfo.UserPrincipalName)) for any questions, inquires, or concerns. `n`nThank you."
	} Else {}

	# Set retention type

If ($shouldRetain -eq "Yes") {
	Get-ADUser $theUser |Move-ADObject -TargetPath "OU=Retained Accounts,DC=sevone,DC=com"
	GAM update user $theUser org "/002 - Retained Accounts"
	} Else {
	Get-ADUser $theUser |Move-ADObject -TargetPath "OU=Former Employees,DC=sevone,DC=com"
	GAM update user $theUser org "/001 - Former Employees"
	Set-ADAccountExpiration $theUser -timespan 180.0:0
	}
	$label_Progress.Text = "Retention & OOO handled..."
	$progressbar_Progress.Value = 50
	$form_Progress.Refresh()

	# Home Drive Archive

$fileToMove = "\\fileserver.sevone.com\File Server\home\$theUser"
$homePath = "\\fileserver.sevone.com\File Server\home"
$homeArchives = "\\fileserver.sevone.com\File Server\home\Former Employee Archives"
$checkExisting = Get-ChildItem -path "\\fileserver.sevone.com\File Server\home\Former Employee Archives" |select-string $theUser\b -Quiet
If (!$checkExisting){ # If Archives does not contain a dupe file of any kind, proceed to copy
	move-item -path "$fileToMove" "$homeArchives"
	}
If ($checkExisting) { # If Archives contains a file, but its not a dupe, ie. (1), proceed to copy
	$checkFile = Get-Item -path "$homeArchives\$theUser*"
	If ($checkFile -match "[0-9]"){ ## Find which number you can use next if dupe with (#)
		$lastFile = $checkFile.Name |findstr "(*)" |select -last 1
		[int]$fileNum = $lastFile -creplace '[^0-9]'
		$fileNum++
		rename-item "$fileToMove" "\\fileserver.sevone.com\File Server\home\$theUser ($fileNum)"
		Move-Item "\\fileserver.sevone.com\File Server\home\$theUser ($fileNum)" "$homeArchives"
		} ElseIf ($checkFile -notmatch "[0-9]") {
			rename-item "$fileToMove" "\\fileserver.sevone.com\File Server\home\$theUser (1)"
			move-item "\\fileserver.sevone.com\File Server\home\$theUser (1)" "$homeArchives"
			}
	}
	$label_Progress.Text = "File Server Home archived..."
	$progressbar_Progress.Value = 60
	$form_Progress.Refresh()

	# Username Archive

$userList = Get-Content "C:\Scripts\Text Files\All Employees List.txt"
$userList |Where-Object {$_ -notlike "$theUser"} |sc "C:\Scripts\Text Files\All Employees List.txt"
Add-Content -Path "C:\Scripts\Text Files\FormerEmployeeUsernamesArchive.txt" "$theUser"
$formerFile = Get-Content "C:\Scripts\Text Files\FormerEmployeeUsernamesArchive.txt" |Sort
$formerFile |Set-Content "C:\SCripts\Text Files\FormerEmployeeUsernamesArchive.txt"

	# Samanage API

$headers = @{
	'Accept' = 'application/XML'
	'Content-Type' = 'text/XML'
	}
$key = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$credPW = cat "C:\Scripts\Text Files\APIcreds.txt" |convertto-securestring -key $key
$creds = New-Object -typename System.Management.Automation.PSCredential -argumentlist "sevoneadmin@sevone.com",$credPW
$GetURI = "https://api.samanage.com/users.xml?email=$theUser@sevone.com"
	$label_Progress.Text = "Samanage disabled..."
	$progressbar_Progress.Value = 70
	$form_Progress.Refresh()
	
	# GET the user - simply used to gather the userID for the PUT request

[string]$XMLResults=(Invoke-WebRequest $GetURI -Credential $creds -Method GET)
$XMLResults |Out-File C:\Temp\XMLResults.xml
$id = Get-Content C:\Temp\XMLResults.xml |Select-String ID |Select -First 1 | % { $_ -creplace '[^0-9]' }
Remove-Item C:\Temp\XMLResults.xml

	# PUT the change

$PutURI = "https://api.samanage.com/users/$id.xml"
[string]$deprovXML = 
	"<user>
	 <role>
	  <id>163943</id>
	  <name>Former Employee</name>
	  </role>
	 <disabled>true</disabled>
	</user>"

Invoke-RestMethod -Body $deprovXML -Credential $creds -Headers $headers -Method PUT -URI $PutURI

	# WebEx API

$XMLURI = "https://sevone.webex.com/WBXService/XMLService"
$ContentType = 'text/xml'
$Method = 'POST'
$Body = (Get-Content "C:\Scripts\Text Files\GetWebExUser.xml").Replace("THE_USERNAME_TO_CHECK","$theUser")
$xmlResponse = Invoke-WebRequest -URI $XMLURI -Body $Body -ContentType $ContentType -Method $Method
If ($xmlResponse -match "FAILURE") {
	# Do nothing because the user does not have a WebEx account
	} Else {
		$Body = (Get-Content "C:\Scripts\Text Files\DelWebExUser.xml").Replace("WEBEX_USERNAME","$theUser")
		Invoke-WebRequest -URI $XMLURI -Method $Method -ContentType $ContentType -Body $Body
		$label_Progress.Text = "WebEx account disabled..."
		$progressbar_Progress.Value = 80
		$form_Progress.Refresh()
		}

	# Jira API

$AuthKey = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$AuthPassword = Get-Content "C:\Scripts\Text Files\JIRA-SVCITAPIAdminCredentials.txt" |ConvertTo-SecureString -Key $AuthKey
$AutoCreds = New-Object System.Management.Automation.PSCredential -ArgumentList "svcitapiadmin",$AuthPassword
$Creds = $AutoCreds.GetNetworkCredential()
$Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $Creds.Username,$Creds.Password)))

$JIRAURI = "https://jira.sevone.com/rest/api/2/user?username=$theUser&expand=groups"
$JIRATestURI = "https://jira-test.sevone.com/rest/api/2/user?username=$theUser&expand=groups"

$checkJIRA = Invoke-RestMethod -URI $JIRAURI -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} |ConvertTo-JSON
$checkJIRATest = Invoke-RestMethod -URI $JIRATestURI -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} |ConvertTo-JSON

If ($checkJIRA) {
	$userToJSON = @{
		"name" = "$theUser"
		}
	$JSON = $userToJSON |ConvertTo-JSON
	$AddURI = "https://jira.sevone.com/rest/api/2/group/user?groupname=inactive"
	Invoke-RestMethod -URI $AddURI -Body $JSON -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} -Method POST -ContentType "application/json"
	$JIRAGroups = $checkJIRA |findstr "@{name=" | % { $_ -creplace '\s' } | % { $_ -creplace '.*groupname=' } | % { $_ -creplace '}"' } | % { $_ -creplace ','} | Where { $_ -ne "inactive" }
	ForEach ($JIRAGroup in $JIRAGroups) {
		$RemoveURI = "https://jira.sevone.com/rest/api/2/group/user?groupname=$JIRAGroup&username=$theUser"
		Invoke-RestMethod -URI $RemoveURI -Body $JSON -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} -Method DELETE -ContentType "application/json"
		}
	Throw-Error "$($theUser) had a JIRA account. Please de-activate."
	}

If ($checkJIRATest) {
	$userToJSON = @{
		"name" = "$theUser"
		}
	$JSON = $userToJSON |ConvertTo-JSON
	$AddURI = "https://jira-test.sevone.com/rest/api/2/group/user?groupname=inactive"
	Invoke-RestMethod -URI $AddURI -Body $JSON -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} -Method POST -ContentType "application/json"
	$JIRATestGroups = $checkJIRATest |findstr "@{name=" | % { $_ -creplace '\s' } | % { $_ -creplace '.*groupname=' } | % { $_ -creplace '}"' } | % { $_ -creplace ','} | Where { $_ -ne "inactive" }
	ForEach ($JIRATestGroup in $JIRATestGroups) {
		$RemoveURI = "https://jira-test.sevone.com/rest/api/2/group/user?groupname=$JIRATestGroup&username=$theUser"
		Invoke-RestMethod -URI $RemoveURI -Body $JSON -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} -Method DELETE -ContentType "application/json"
		}
	Throw-Error "$($theUser) had a JIRA-Test account. Please de-activate."
	}
	
$checkConfluence = Get-Content "C:\Scripts\Text Files\ConfluenceLocalAccounts.txt" |Select-String $theUser -Quiet
If ($checkConfluence -eq $True) { Throw-Error "$($theUser) has a local Confluence account. Please de-provision manually." }

# CUCM API :(
# Zendesk API :(
# Check SVC Accounts

# Logging

$date = Get-Date
Echo "$date :: $loggedInUser deprovisioned $theUser" >> $logFile

	$label_Progress.Text = "Done!"
	$progressbar_Progress.Value = 100
	$form_Progress.Refresh()
	Sleep 3

$form_Progress.Dispose()
Run-MainMenu
}

Start-Form