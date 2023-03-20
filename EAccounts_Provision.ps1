# Employment Tool v3.0
# 	Employee Accounts - Provision

	# Gather information for provisioning
Function ClearBlankSpaces ($text,$name){
If ($text -match '^.* ') {
	$($name).Text = $($name).Text -creplace "\s"
	$($name).SelectionStart = $($name).Text.Length + 1
	}
}

Function Start-Form {

$ADList = Get-ADUser -Filter *

$form_StartForm = New-Form
$form_StartForm.Size = Set-Size 420 490
$form_StartForm.Text = "  Employment Tool - Provisioning"
$form_StartForm.SizeGripStyle = 'Hide'
$form_StartForm.StartPosition = 'CenterScreen'
$form_StartForm.FormBorderStyle = 'Fixed3D'
$form_StartForm.Font = "Arial,10,style=Bold"
$form_StartForm.TopMost = $True
$form_StartForm.KeyPreview = $True
$form_StartForm.MaximizeBox = $False
$form_StartForm.MinimizeBox = $False
$form_StartForm.Icon = Set-Icon

$label_FirstName = New-Label
$label_FirstName.Size = Set-Size 140 20
$label_FirstName.Location = Set-Point 10 10
$label_FirstName.Text = "First Name"
    $form_StartForm.Controls.Add($label_FirstName)

$textBox_FirstName = New-TextBox
$textBox_FirstName.Size = Set-Size 240 20
$textBox_FirstName.Location = Set-Point 160 10
$textBox_FirstName.Add_TextChanged({
	ClearBlankSpaces $textBox_FirstName.Text '$textBox_FirstName'
	Validate-Form
	})
    $form_StartForm.Controls.Add($textBox_FirstName)

$label_LastName = New-Label
$label_LastName.Size = Set-Size 140 20
$label_LastName.Location = Set-Point 10 40
$label_LastName.Text = "Last Name"
	$form_StartForm.Controls.Add($label_LastName)

$textBox_LastName = New-TextBox
$textBox_LastName.Size = Set-Size 240 20
$textBox_LastName.Location = Set-Point 160 40
$textBox_LastName.Add_TextChanged({Validate-Form})
    $form_StartForm.Controls.Add($textBox_LastName)

$label_UserName = New-Label
$label_UserName.Size = Set-Size 140 20
$label_UserName.Location = Set-Point 10 70
$label_UserName.Text = "User Name"
    $form_StartForm.Controls.Add($label_UserName)

$textBox_UserName = New-TextBox
$textBox_UserName.Size = Set-Size 240 20
$textBox_UserName.Location = Set-Point 160 70
$textBox_UserName.MaxLength	= 20
$textBox_UserName.Add_TextChanged({
	$script:theUser = $textBox_UserName.Text
	$userNameLength	= $textBox_UserName.Text |Measure-Object -Character
	If ($userNameLength.Characters -eq 20) {
		Throw-Error "Username can not exceed 20 characters."
		}
	$checkAD = $ADList.SamAccountName |Select-String "^$($theUser)\b$" -Quiet
	$checkArchives = GC "C:\Scripts\Text Files\FormerEmployeeUsernamesArchive.txt" |Select-String "^$($theUser)\b$" -Quiet
	If ($checkAD -eq $True -OR $checkArchives -eq $True -OR $theUser -eq "") {
		$label_UserName.ForeColor = "Red"
		$script:usernameValid = "False"
		Validate-Form
		} Else {
			$label_UserName.ForeColor = "Green"
			$script:usernameValid = "True"
			Validate-Form
			}
	})
    $form_StartForm.Controls.Add($textBox_UserName)

$label_Manager = New-Label
$label_Manager.Size = Set-Size 140 20
$label_Manager.Location = Set-Point 10 100
$label_Manager.Text = "Manager"
    $form_StartForm.Controls.Add($label_Manager)

$textbox_Manager = New-TextBox
$textbox_Manager.Size = Set-Size 240 20
$textbox_Manager.Location = Set-Point 160 100
$textbox_Manager.Add_TextChanged({
	$checkAD = $ADList.SamAccountName |Select-String "^$($textbox_Manager.Text)\b$" -Quiet
	If ($checkAD -eq $True) {
		$label_Manager.ForeColor = "Green"
		$script:managerCorrect = "True"
		Validate-Form
		}
	Else {
		$label_Manager.ForeColor = "Red"
		$script:managerCorrect = "False"
		Validate-Form
		}
	})
	$form_StartForm.Controls.Add($textbox_Manager)

$label_Title = New-Label
$label_Title.Size = Set-Size 140 20
$label_Title.Location = Set-Point 10 130
$label_Title.Text = "Title"
    $form_StartForm.Controls.Add($label_Title)

$textBox_Title = New-TextBox
$textBox_Title.Size = Set-Size 240 20
$textBox_Title.Location = Set-Point 160 130
$textBox_Title.Add_TextChanged({Validate-Form})
    $form_StartForm.Controls.Add($textBox_Title)

$label_Dept = New-Label
$label_Dept.Size = Set-Size 140 20
$label_Dept.Location = Set-Point 10 160
$label_Dept.Text = "Department"
    $form_StartForm.Controls.Add($label_Dept)

$comboBox_Dept = New-ComboBox
$comboBox_Dept.Size = Set-Size 240 20
$comboBox_Dept.Location = Set-Point 160 160
$comboBox_Dept.DropDownStyle = "DropDownList"
$comboBox_Dept.Add_SelectedIndexChanged({Validate-Form})
$comboBox_Dept_List = Get-Content "C:\Scripts\Text Files\DeptCodes.txt"
ForEach ($dept in $comboBox_Dept_List) {
    $comboBox_Dept.Items.Add($dept)
	}
    $form_StartForm.Controls.Add($comboBox_Dept)

$label_OfficeLoc = New-Label
$label_OfficeLoc.Size = Set-Size 140 20
$label_OfficeLoc.Location = Set-Point 10 190
$label_OfficeLoc.Text = "Office Location"
    $form_StartForm.Controls.Add($label_OfficeLoc)

$comboBox_OfficeLoc = New-ComboBox
$comboBox_OfficeLoc.Size = Set-Size 240 20
$comboBox_OfficeLoc.Location = Set-Point 160 190
$comboBox_OfficeLoc.DropDownStyle = "DropDownList"
$comboBox_OfficeLoc.Add_SelectedIndexChanged({Validate-Form})
$comboBox_OfficeLoc_List = Get-Content "C:\Scripts\Text Files\Offices.txt"
ForEach ($office in $comboBox_OfficeLoc_List) {
    $comboBox_OfficeLoc.Items.Add($office)
	}
    $form_StartForm.Controls.Add($comboBox_OfficeLoc)

$label_CorpGroup = New-Label
$label_CorpGroup.Size = Set-Size 140 20
$label_CorpGroup.Location = Set-Point 10 220
$label_CorpGroup.Text = "Corporate Group"
    $form_StartForm.Controls.Add($label_CorpGroup)

$comboBox_corpGroup = New-ComboBox
$comboBox_corpGroup.Size = Set-Size 240 20
$comboBox_corpGroup.Location = Set-Point 160 220
$comboBox_corpGroup.DropDownStyle = "DropDownList"
$comboBox_corpGroup.Add_SelectedIndexChanged({Validate-Form})
$comboBox_CorpGroup_List = Import-CSV "C:\Scripts\Text Files\GoogleGroups.csv" |Where { $_.Type -eq "corporate"} |Sort name
ForEach ($corpGroup in $comboBox_CorpGroup_List) {
    $comboBox_corpGroup.Items.Add($corpGroup.name)
	}
    $form_StartForm.Controls.Add($comboBox_corpGroup)

$label_LocGroup = New-Label
$label_LocGroup.Size = Set-Size 140 20
$label_LocGroup.Location = Set-Point 10 250
$label_LocGroup.Text = "Location Group"
    $form_StartForm.Controls.Add($label_LocGroup)

$comboBox_LocGroup = New-ComboBox
$comboBox_LocGroup.Size = Set-Size 240 20
$comboBox_LocGroup.Location = Set-Point 160 250
$comboBox_LocGroup.DropDownStyle = "DropDownList"
$comboBox_LocGroup.Add_SelectedIndexChanged({Validate-Form})
$comboBox_LocGroup_List                		= Import-CSV "C:\Scripts\Text Files\GoogleGRoups.csv" |Where { $_.Type -eq "location"} |Sort name
ForEach ($locGroup in $comboBox_LocGroup_List) {
    $comboBox_LocGroup.Items.Add($locGroup.name)
	}
    $form_StartForm.Controls.Add($comboBox_LocGroup)

$label_DeptGroup = New-Label
$label_DeptGroup.Size = Set-Size 140 20
$label_DeptGroup.Location = Set-Point 10 280
$label_DeptGroup.Text = "Department Group"
    $form_StartForm.Controls.Add($label_DeptGroup)

$comboBox_DeptGroup = New-ComboBox
$comboBox_DeptGroup.Size = Set-Size 240 20
$comboBox_DeptGroup.Location = Set-Point 160 280
$comboBox_DeptGroup.DropDownStyle = "DropDownList"
$comboBox_DeptGroup.Add_SelectedIndexChanged({Validate-Form})
$comboBox_DeptGroup_List = Import-CSV "C:\Scripts\Text Files\GoogleGroups.csv" |Where { $_.Type -eq "department"} |Sort name
ForEach ($deptGroup in $comboBox_DeptGroup_List) {
    $comboBox_DeptGroup.Items.Add($deptGroup.name)
	}
    $form_StartForm.Controls.Add($comboBox_DeptGroup)

$label_EmployeeType = New-Label
$label_EmployeeType.Size = Set-Size 140 20
$label_EmployeeType.Location = Set-Point 10 310
$label_EmployeeType.Text = "Employee Type"
    $form_StartForm.Controls.Add($label_EmployeeType)

$comboBox_EmployeeType = New-ComboBox
$comboBox_EmployeeType.Size = Set-Size 240 20
$comboBox_EmployeeType.Location = Set-Point 160 310
$comboBox_EmployeeType.DropDownStyle = "DropDownList"
$comboBox_EmployeeType.Add_SelectedIndexChanged({Validate-Form})
$comboBox_EmployeeType_List =
    "Employee",
	"Intern",
    "Programista",
	"Jeavio",
    "Contractor",
    "Sub-Contractor"
ForEach ($employeeType in $comboBox_EmployeeType_List) {
    $comboBox_EmployeeType.Items.Add($employeeType)
	}
    $form_StartForm.Controls.Add($comboBox_EmployeeType)

$expirationDate_Label = New-Label
$expirationDate_Label.Size = Set-Size 140 20
$expirationDate_Label.Location = Set-Point 10 340
$expirationDate_Label.Text = "Expiration Date"
    $form_StartForm.Controls.Add($expirationDate_Label)

$dateTimePicker_ExpirationDate = New-DatePicker
$dateTimePicker_ExpirationDate.Size = Set-Size 240 20
$dateTimePicker_ExpirationDate.Location = Set-Point 160 340
$dateTimePicker_ExpirationDate.Enabled = $False
    $form_StartForm.Controls.Add($dateTimePicker_ExpirationDate)

$checkBox_AccountExpires = New-CheckBox
$checkBox_AccountExpires.Size = Set-Size 140 20
$checkBox_AccountExpires.Location = Set-Point 10 370
$checkBox_AccountExpires.Text = "Account Expires"
$checkBox_AccountExpires.Add_CheckStateChanged({
	If ($checkBox_AccountExpires.CheckState -eq $True) {
		$dateTimePicker_ExpirationDate.Enabled = $True
		} Else {
			$dateTimePicker_ExpirationDate.Enabled = $False
			}
	})
	$form_StartForm.Controls.Add($checkBox_AccountExpires)

$checkBox_IsManager = New-CheckBox
$checkBox_IsManager.Size = Set-Size 140 20
$checkBox_IsManager.Location = Set-Point 150 370
$checkBox_IsManager.Text = "Is Manager"
    $form_StartForm.Controls.Add($checkBox_IsManager)

$checkBox_WebEx = New-CheckBox
$checkBox_WebEx.Size = Set-Size 140 20
$checkBox_WebEx.Location = Set-Point 10 390
$checkBox_WebEx.Text = "WebEx"
	$form_StartForm.Controls.Add($checkBox_WebEx)

$checkBox_JIRA = New-CheckBox
$checkBox_JIRA.Size = Set-Size 140 20
$checkBox_JIRA.Location = Set-Point 150 390
$checkBox_JIRA.Text = "JIRA"
	$form_StartForm.Controls.Add($checkBox_JIRA)

$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 270 420
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form_StartForm.Controls.Add($button_Cancel)
    $form_StartForm.CancelButton = $button_Cancel

$button_OK = New-Button
$button_OK.Size = Set-Size 60 30
$button_OK.Location = Set-Point 340 420
$button_OK.Text = "OK"
$button_OK.Enabled = $False
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button_OK.Add_Click({
    $script:theFirstName = $textBox_FirstName.Text
    $script:theLastName = $textBox_LastName.Text
    $script:theDisplayName = "$($textBox_LastName.Text), $($textBox_FirstName.Text)"
    $script:theEmail = "$theUser@domain.com"
	$script:theManager = $textbox_Manager.Text
    $script:theTitle = $textBox_Title.Text
    $script:theDept = $comboBox_Dept.SelectedItem
	$script:theOfficeLoc = $comboBox_OfficeLoc.SelectedItem
	$script:theEmployeeType = $comboBox_EmployeeType.SelectedItem
    If ($comboBox_CorpGroup.SelectedIndex -gt -1) {$script:theCorpGroup = $comboBox_CorpGroup.SelectedItem} Else {$script:theCorpGroup = "N/A"}
    If ($comboBox_LocGroup.SelectedIndex -gt -1) {$script:theLocGroup = $comboBox_LocGroup.SelectedItem} Else {$script:theLocGroup = "N/A"}
    If ($comboBox_DeptGroup.SelectedIndex -gt -1) {$script:theDeptGroup = $comboBox_DeptGroup.SelectedItem} Else {$script:theDeptGroup = "N/A"}
    If ($checkbox_IsManager.CheckState -eq "Checked") {$script:theManagerStatus = "True"}
    If ($checkBox_IsManager.CheckState -eq "Unchecked") {$script:theManagerStatus = "False"}
	If ($checkBox_WebEx.CheckState -eq "Checked") {$script:theWebExStatus = "True"}
	If ($checkBox_WebEx.CheckState -eq "Unchecked") {$script:theWebExStatus = "False"}
	If ($checkBox_JIRA.CheckState -eq "Checked") {$script:theJIRAStatus = "True"}
	If ($checkBox_JIRA.CheckState -eq "Unchecked") {$script:theJIRAStatus = "False"}
    If ($dateTimePicker_ExpirationDate.Enabled -eq $True) {[dateTime]$script:theExpirationDate = $dateTimePicker_ExpirationDate.Value.ToString("MM/dd/yyyy") } Else {$script:theExpirationDate = "Never"}
	})
    $form_StartForm.Controls.Add($button_OK)
    $form_StartForm.AcceptButton = $button_OK

$fr = $form_StartForm.ShowDialog()
If ($fr -eq "Cancel") { $form_StartForm.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") { $form_StartForm.Dispose() ; Confirmation-Screen }
}

	# Enable the okay button only once all fields have values

Function Validate-Form {
If (
    $textBox_FirstName.Text -AND
    $textBox_LastName.Text -AND
    $textBox_UserName.Text -AND
	$usernameValid -eq "True" -AND
	$managerCorrect -eq "True" -AND
    $textBox_Title.Text -AND
    $comboBox_Dept.SelectedIndexChanged -gt -1 -AND
    $comboBox_OfficeLoc.SelectedIndex -gt -1 -AND
    $comboBox_EmployeeType.SelectedIndex -gt -1
) { $button_OK.Enabled = $True } Else { $button_OK.Enabled = $False }
}

	# Confirm provisioning

Function Confirmation-Screen {
$form_ConfirmationScreen = New-Form
$form_ConfirmationScreen.Text = "  Employment Tool"
$form_ConfirmationScreen.AutoSize = $True
$form_ConfirmationScreen.SizeGripStyle = 'Hide'
$form_ConfirmationScreen.StartPosition = 'CenterScreen'
$form_ConfirmationScreen.FormBorderStyle = 'Fixed3D'
$form_ConfirmationScreen.Font = "Arial,10,style=Bold"
$form_ConfirmationScreen.TopMost = $True
$form_ConfirmationScreen.KeyPreview = $True
$form_ConfirmationScreen.MaximizeBox = $False
$form_ConfirmationScreen.MinimizeBox = $False
$form_ConfirmationScreen.Icon = Set-Icon

$label_theFirstName = New-Label
$label_theFirstName.AutoSize = $True
$label_theFirstName.Location = Set-Point 10 10
$label_theFirstName.Text = "First Name: $theFirstName"
    $form_ConfirmationScreen.Controls.Add($label_theFirstName)

$label_theLastName = New-Label
$label_theLastName.AutoSize = $True
$label_theLastName.Location = Set-Point 10 40
$label_theLastName.Text = "Last Name: $theLastName"
    $form_ConfirmationScreen.Controls.Add($label_theLastName)

$label_theUserName = New-Label
$label_theUserName.AutoSize = $True
$label_theUserName.Location = Set-Point 10 70
$label_theUserName.Text = "Username: $theUser"
    $form_ConfirmationScreen.Controls.Add($label_theUserName)

$label_theManager = New-Label
$label_theManager.AutoSize = $True
$label_theManager.Location = Set-Point 10 100
$label_theManager.Text = "Manager: $theManager"
    $form_ConfirmationScreen.Controls.Add($label_theManager)

$label_theTitle = New-Label
$label_theTitle.AutoSize = $True
$label_theTitle.Location = Set-Point 10 130
$label_theTitle.Text = "Title: $theTitle"
    $form_ConfirmationScreen.Controls.Add($label_theTitle)

$label_theDept = New-Label
$label_theDept.AutoSize = $True
$label_theDept.Location = Set-Point 10 160
$label_theDept.Text = "Department: $theDept"
    $form_ConfirmationScreen.Controls.Add($label_theDept)

$label_theOfficeLoc = New-Label
$label_theOfficeLoc.AutoSize = $True
$label_theOfficeLoc.Location = Set-Point 10 190
$label_theOfficeLoc.Text = "Office: $theOfficeLoc"
    $form_ConfirmationScreen.Controls.Add($label_theOfficeLoc)

$label_theCorpGroup = New-Label
$label_theCorpGroup.AutoSize = $True
$label_theCorpGroup.Location = Set-Point 10 220
$label_theCorpGroup.Text = "Corporate Group: $theCorpGroup"
    $form_ConfirmationScreen.Controls.Add($label_theCorpGroup)

$label_theLocGroup = New-Label
$label_theLocGroup.AutoSize = $True
$label_theLocGroup.Location = Set-Point 10 250
$label_theLocGroup.Text = "Location Group: $theLocGroup"
    $form_ConfirmationScreen.Controls.Add($label_theLocGroup)

$label_theDeptGroup = New-Label
$label_theDeptGroup.AutoSize = $True
$label_theDeptGroup.Location = Set-Point 10 280
$label_theDeptGroup.Text = "Dept Group: $theDeptGroup"
    $form_ConfirmationScreen.Controls.Add($label_theDeptGroup)

$label_theEmployeeType = New-Label
$label_theEmployeeType.AutoSize = $True
$label_theEmployeeType.Location = Set-Point 10 310
$label_theEmployeeType.Text = "Employee Type: $theEmployeeType"
    $form_ConfirmationScreen.Controls.Add($label_theEmployeeType)

$label_theManagerStatus = New-Label
$label_theManagerStatus.AutoSize = $True
$label_theManagerStatus.Location = Set-Point 10 340
$label_theManagerStatus.Text = "Is Manager? $theManagerStatus"
    $form_ConfirmationScreen.Controls.Add($label_theManagerStatus)
	
$label_theWebExStatus = New-Label
$label_theWebExStatus.Autosize = $True
$label_theWebExStatus.Location = Set-Point 10 370
$label_theWebExStatus.Text = "WebEx? $theWebExStatus"
	$form_ConfirmationScreen.Controls.Add($label_theWebExStatus)
	
$label_theJIRAStatus = New-Label
$label_theJIRAStatus.Autosize = $True
$label_theJIRAStatus.Location = Set-Point 10 400
$label_theJIRAStatus.Text = "JIRA? $theJIRAStatus"
	$form_ConfirmationScreen.Controls.Add($label_theJIRAStatus)

$label_theExpirationDate = New-Label
$label_theExpirationDate.AutoSize = $True
$label_theExpirationDate.Location = Set-Point 10 430
$label_theExpirationDate.Text = "Account Expires: $theExpirationDate"
    $form_ConfirmationScreen.Controls.Add($label_theExpirationDate)

$label_Warning = New-Label
$label_Warning.Location = Set-Point 60 460
$label_warning.AutoSize = $True
$label_Warning.ForeColor = "Red"
$label_Warning.Font = "Arial,8,style=Italic"
$label_Warning.Text = "* By clicking OK you will begin to provision this user!"
    $form_ConfirmationScreen.Controls.add($label_Warning)

$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 270 490
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form_ConfirmationScreen.Controls.Add($button_Cancel)
    $form_ConfirmationScreen.CancelButton = $button_Cancel

$button_OK = New-Button
$button_OK.Size = Set-Size 60 30
$button_OK.Location = Set-Point 340 490
$button_OK.Text = "OK"
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form_ConfirmationScreen.Controls.Add($button_OK)
    $form_ConfirmationScreen.AcceptButton = $button_OK

$fr = $form_ConfirmationScreen.ShowDialog()
If ($fr -eq "Cancel") { $form_ConfirmationScreen.Dispose() ; Start-Form }
If ($fr -eq "OK") { Begin-Provisioning }
}

	# Begin provisioning

Function Begin-Provisioning {
$thePassword = (ConvertTo-SecureString -AsPlainText "Day1@domain!" -Force)
$theOU = "CN=Users,DC=domain,DC=com"
$theDescription = "$theEmployeeType - $theDept"

	# Set up progress form
	
$form_BeginProvisioning	= New-Form
$form_BeginProvisioning.Size = Set-Size 290 140
$form_BeginProvisioning.Text = "  Employment Tool"
$form_BeginProvisioning.SizeGripStyle = 'Hide'
$form_BeginProvisioning.StartPosition = 'CenterScreen'
$form_BeginProvisioning.FormBorderStyle = 'Fixed3D'
$form_BeginProvisioning.Font = "Arial,10,style=Bold"
$form_BeginProvisioning.TopMost = $True
$form_BeginProvisioning.KeyPreview = $True
$form_BeginProvisioning.MaximizeBox = $False
$form_BeginProvisioning.MinimizeBox	= $False
$form_BeginProvisioning.Icon = Set-Icon

$label = New-Label
$label.Size = Set-Size 250 20
$label.Location = Set-Point 10 10
$label.Text = "Preparing to provision..."
    $form_BeginProvisioning.Controls.Add($label)

$progression = New-ProgressBar
$progression.Size = Set-Size 250 20
$progression.Location = Set-Point 10 40
$progression.Value = 0
    $form_BeginProvisioning.Controls.Add($progression)

$form_BeginProvisioning.Show()
	
	# Create User
	
	$label.Text = "Creating User..."
	$progression.Value = 5
	$form_BeginProvisioning.Refresh()

	# Create the account.. This tossed an error for some unknown reason at times, but unsure why...
	
New-ADUser "$theDisplayName" -SamAccountName $theUser -UserPrincipalName $theEmail

	# Finalize User Properties
	
	$label.Text = "Setting Account Information..."
	$progression.Value = 20
	$form_BeginProvisioning.Refresh()

Set-ADUser $theUser -Email $theEmail
Set-ADUser $theUser -GivenName $theFirstName
Set-ADUser $theUser -Surname $theLastName
Set-ADAccountPassword $theUser -NewPassword $thePassword
Set-ADUser $theUser -ChangePasswordAtLogon $False
Set-ADUser $theUser -Enabled $True
Set-ADUser $theUser -DisplayName "$theDisplayName"
Set-ADUser $theUser -Company "domain, Inc."
Set-ADUser $theUser -PasswordNeverExpires $True
Set-ADUser $theUser -Department $theDept
Set-ADUser $theUser -Title $theTitle
Set-ADUser $theUser -Office $theOfficeLoc
Set-ADUser $theUser -Manager $theManager
Set-ADUser $theUser -Description $theDescription
If ($theExpirationDate -ne "Never") {
	Set-ADAccountExpiration $theUser $theExpirationDate.AddDays(1)
	}
	
	# Add AD Groups
	
	$label.Text = "Setting AD Groups..."
	$progression.Value = 30
	$form_BeginProvisioning.Refresh()

Add-ADGroupMember domain $theUser
Add-ADGroupMember NonAdmins $theUser
Add-ADGroupMember SSLVPN $theUser
Add-ADGroupMember Wireless $theUser
If ($theEmployeeType -match "Programista") {Add-ADGroupMember "Programista" $theUser}
If ($theEmployeeType -match "Jeavio") {Add-ADGroupMember "Jeavio" $theUser}
If ($theEmployeeType -match "Contractor" -OR $theEmployeeType -match "Sub-Contractor") {Add-ADGroupMember "Contractors" $theUser}
If ($theEmployeeType -match "Employee" -OR $theEmployeeType -match "Intern") {Add-ADGroupMember "domain Staff" $theUser} # Dictates domain Source access

	# Set Home Drive
	
	$label.Text = "Setting Home Drive..."
	$progression.Value = 40
	$form_BeginProvisioning.Refresh()

$homeDriveCheck	= Test-Path "\\fileserver.domain.com\File Server\home\$theUser"
If (!$homeDriveCheck) {
	New-Item -path "\\fileserver.domain.com\File Server\home\$theUser" -type Directory
	sleep 2
	$ACL = Get-ACL "\\fileserver.domain.com\File Server\home\$theUser"
	$AR = New-Object system.security.accesscontrol.filesystemaccessrule("$theUser","FullControl","ContainerInherit,ObjectInherit","None","Allow")
	$ACL.SetAccessRule($AR)
	Set-ACL "\\fileserver.domain.com\File Server\home\$theUser" $ACL
	Set-ADUser $theUser -HomeDrive "Z:" -HomeDirectory "\\fileserver.domain.com\File Server\home\$theUser"
	}

	# Sync Google 50% mark
	
	$label.Text = "Synchronizing Data to Google..."
	$progression.Value = 50
	$form_BeginProvisioning.Refresh()
	
Sync-Google -Wait

	# Google OU & Groups

	$label.Text = "Setting Google OU & Groups..."
	$progression.Value = 80
	$form_BeginProvisioning.Refresh()
	
GAM update user $theUser org "/005 - 2FA Exception"
If ($theDeptGroup -eq "N/A") {} Else {GAM update group $theDeptGroup add user $theUser}
If ($theCorpGroup -eq "N/A") {} Else {GAM update group $theCorpGroup add user $theUser}
If ($theLocGroup -ne "N/A (International)" -OR $theLocGroup -ne "N/A") {GAM update group $theLocGroup add user $theUser}
If ($theManagerStatus -eq "True") {GAM update group managers add user $theUser}

	# WebEx API
Function OldWebExAPI {
If ($theWebExStatus -eq "True") {
	$URI = 'https://domain.webex.com/WBXService/XMLService'
	$ContentType = 'text/xml'
	$Method = 'POST'
	
	$CheckExistingBody = (Get-Content "C:\Scripts\Text Files\GetWebExUser.xml").Replace("THE_USERNAME_TO_CHECK","$theUser")
	
	$checkExistingResults = (Invoke-WebRequest -URI $URI -Method $Method -ContentType $ContentType -Body $CheckExistingBody).Content

If ($checkExistingResults -match "FAILURE") { # Proceed to provision because there is no user by that name
	$BodyFixFirst = (Get-Content "C:\Scripts\Text Files\CreateWebExUser.xml").Replace("THE_FIRSTNAME","$theFirstName")
	$BodyFixLast = $BodyFixFirst.Replace("THE_LASTNAME","$theLastName")
	$BodyFixUsername = $BodyFixLast.Replace("THE_USERNAME","$theUser")
	$ProvisionBody = $BodyFixUsername.Replace("THE_EMAIL","$theEmail")
	Invoke-WebRequest -URI $URI -Method $Method -ContentType $ContentType -Body $ProvisionBody
	} Else { # Throw-Error advising username already used
		Throw-Error "The username $($theUser) already exists in the WebEx System!"
		}
	}
}

# Code moved to the end of block

	# JIRA & JIRA-Test API

If ($theJIRAStatus -eq "True") {
	$JIRAURI = "https://jira.domain.com/rest/api/2/user/search?username=$theUser"
	$JIRATestURI = "https://jira-test.domain.com/rest/api/2/user/search?username=$theUser"
	$AuthKey = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
	$AuthPassword = Get-Content "C:\Scripts\Text Files\JIRA-SVCITAPIAdminCredentials.txt" |ConvertTo-SecureString -Key $AuthKey
	$AutoCreds = New-Object System.Management.Automation.PSCredential -ArgumentList "svcitapiadmin",$AuthPassword
	$Creds = $AutoCreds.GetNetworkCredential()
	$Base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $Creds.Username,$Creds.Password)))
	Sleep 1
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
			Invoke-RestMethod -URI $JIRAGroupURI -Headers @{"Authorization"=("Basic {0}" -f $Base64AuthInfo)} -Method POST -ContentType "application/json" -Body $JSON
		}
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
	}

	# Homebrew
	
	$label.Text = "Adding to Homebrew..."
	$progression.Value = 90
	$form_BeginProvisioning.Refresh()

Add-Content -Path "C:\Scripts\Text Files\All Employees List.txt" "$theUser"
$homebrewFile = Get-Content "C:\Scripts\Text Files\All Employees List.txt" |Sort
$homebrewFile |Set-Content "C:\Scripts\Text Files\All Employees List.txt"

	# Logging
	
	$label.Text = "Logging Creation..."
	$progression.Value = 95
	$form_BeginProvisioning.Refresh()

$date	= Get-Date
Write-Log "$loggedInUser provisioned an account for $theFirstName $theLastName ($theEmail)."

	#DONE!
	$label.Text = "Done!"
	$progression.Value = 100
	$form_BeginProvisioning.Refresh()

Sleep 3
If ($theWebExStatus -eq $True) {
	Run-CreateWebExAccount
	$form_BeginProvisioning.Dispose()
	} Else {
		$form_BeginProvisioning.Dispose()
		Run-MainMenu
		}
}

Start-Form
