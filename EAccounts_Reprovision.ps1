# Employment Tool v3.0
# 	Employee Accounts - Re-Provision

	# Gather information for re-provisioning

Function Start-Form {
$form_StartForm = New-Form
$form_StartForm.Text = "  Employment Tool - ReProvisioning"
$form_StartForm.Size = Set-Size 420 410
$form_StartForm.SizeGripStyle = 'Hide'
$form_StartForm.StartPosition = 'CenterScreen'
$form_StartForm.FormBorderStyle = 'Fixed3D'
$form_StartForm.Font = "Arial,10,style=Bold"
$form_StartForm.TopMost = $True
$form_StartForm.KeyPreview = $True
$form_StartForm.MaximizeBox = $False
$form_StartForm.MinimizeBox = $False
$form_StartForm.Icon = Set-Icon

$label_Name = New-Label
$label_Name.Size = Set-Size 140 20
$label_Name.Location = Set-Point 10 10
$label_Name.Text = "Employee Name"
	$form_StartForm.Controls.Add($label_Name)

$comboBox_Name = New-ComboBox
$comboBox_Name.Size = Set-Size 240 20
$comboBox_Name.Location = Set-Point 160 10
$comboBox_Name.DropDownStyle = "DropDownList"
$comboBox_Name_List1 = Get-ADUser -Filter * -SearchBase "OU=Former Employees,DC=sevone,DC=com" |Select SamAccountName |Sort SamAccountName
$comboBox_Name_List2 = Get-ADUser -Filter * -SearchBase "OU=Retained Accounts,DC=sevone,DC=com" |Select SamAccountName |Sort SamAccountName
$comboBox_Name_List = $comboBox_Name_List1  + $comboBox_Name_List2 |Sort SamAccountName
$comboBox_Name.Add_SelectedIndexChanged({Check-Fill})
ForEach ($user in $comboBox_Name_List) {
     $comboBox_Name.Items.Add($user.SamAccountName)
}
    $form_StartForm.Controls.Add($comboBox_Name)

$label_Manager = New-Label
$label_Manager.Size = Set-Size 140 20
$label_Manager.Location = Set-Point 10 40
$label_Manager.Text = "Manager"
    $form_StartForm.Controls.Add($label_Manager)

$ADList = Get-ADUser -Filter * |Select SamAccountName
$textbox_Manager = New-Textbox
$textbox_Manager.Size = Set-Size 240 20
$textbox_Manager.Location = Set-Point 160 40
$textbox_Manager.Add_TextChanged({
	$checkAD = $ADList.SamAccountName |Select-String "^$($textbox_Manager.Text)\b$" -Quiet
	If ($checkAD -eq $True) {
		$label_Manager.ForeColor = "Green"
		$script:managerCorrect = "True"
		Check-Fill
	} Else {
		$label_Manager.ForeColor = "Red"
		$script:managerCorrect = "False"
		Check-Fill
	}
})
	$form_StartForm.Controls.Add($textbox_Manager)

$label_Title = New-Label
$label_Title.Size = Set-Size 140 20
$label_Title.Location = Set-Point 10 70
$label_Title.Text = "Title"
    $form_StartForm.Controls.Add($label_Title)

$textBox_Title = New-TextBox
$textBox_Title.Size = Set-Size 240 20
$textBox_Title.Location = Set-Point 160 70
$textBox_Title.Add_TextChanged({Check-Fill})
	$form_StartForm.Controls.Add($textBox_Title)

$label_Department = New-Label
$label_Department.Size = Set-Size 140 20
$label_Department.Location = Set-Point 10 100
$label_Department.Text = "Department"
	$form_StartForm.Controls.Add($label_Department)

$comboBox_Department = New-ComboBox
$comboBox_Department.DropDownStyle 	= "DropDownList"
$comboBox_Department.Size = Set-Size 240 20
$comboBox_Department.Location = Set-Point 160 100
$comboBox_Department.Add_SelectedIndexChanged({Check-Fill})
$deptCodes = Get-Content "C:\Scripts\Text Files\DeptCodes.txt"
ForEach ($dept in $deptCodes) {
	$comboBox_Department.Items.Add($dept)
}
	$form_StartForm.Controls.Add($comboBox_Department)

$label_OfficeLoc = New-Label
$label_OfficeLoc.Size = Set-Size 140 20
$label_OfficeLoc.Location = Set-Point 10 130
$label_OfficeLoc.Text = "Office Location"
	$form_StartForm.Controls.Add($label_OfficeLoc)

$comboBox_OfficeLoc = New-ComboBox
$comboBox_OfficeLoc.Size = Set-Size 240 20
$comboBox_OfficeLoc.Location = Set-Point 160 130
$comboBox_OfficeLoc.DropDownStyle = "DropDownList"
$comboBox_OfficeLoc.Add_SelectedIndexChanged({Check-Fill})
$comboBox_OfficeLoc_Choices = Get-Content "C:\Scripts\Text Files\Offices.txt"
ForEach ($choice in $comboBox_OfficeLoc_Choices) {
	$comboBox_OfficeLoc.Items.Add($choice)
}
	$form_StartForm.Controls.Add($comboBox_OfficeLoc)

$label_CorpGroup = New-Label
$label_CorpGroup.Size = Set-Size 140 20
$label_CorpGroup.Location  = Set-Point 10 160
$label_CorpGroup.Text = "Corporate Group"
	$form_StartForm.Controls.Add($label_CorpGroup)

$comboBox_CorpGroup = New-ComboBox
$comboBox_CorpGroup.Size = Set-Size 240 20
$comboBox_CorpGroup.Location = Set-Point 160 160
$comboBox_CorpGroup.DropDownStyle = "DropDownList"
$comboBox_CorpGroup.Add_SelectedIndexChanged({Check-Fill})
$comboBox_CorpGroup_Choices = Import-CSV "C:\Scripts\Text Files\GoogleGroups.csv" |Where { $_.Type -eq "corporate" }
ForEach ($corpGroup in $comboBox_CorpGroup_Choices) {
    $comboBox_CorpGroup.Items.Add($corpGroup.name)
}
    $form_StartForm.Controls.Add($comboBox_CorpGroup)

$label_LocGroup = New-Label
$label_LocGroup.Size = Set-Size 140 20
$label_LocGroup.Location = Set-Point 10 190
$label_LocGroup.Text = "Location Group"
    $form_StartForm.Controls.Add($label_LocGroup)

$comboBox_LocGroup = New-ComboBox
$comboBox_LocGroup.Size = Set-Size 240 20
$comboBox_LocGroup.Location = Set-Point 160 190
$comboBox_LocGroup.DropDownStyle = "DropDownList"
$comboBox_LocGroup.Add_SelectedIndexChanged({Check-Fill})
$comboBox_LocGroup_Choices = Import-CSV "C:\Scripts\Text Files\GoogleGroups.csv" |Where { $_.Type -eq "location" } |Sort name
ForEach ($locGroup in $comboBox_LocGroup_Choices) {
    $comboBox_LocGroup.Items.Add($locGroup.name)
}
    $form_StartForm.Controls.Add($comboBox_LocGroup)

$label_DeptGroup = New-Label
$label_DeptGroup.Size = Set-Size 140 20
$label_DeptGroup.Location  = Set-Point 10 220
$label_DeptGroup.Text = "Dept Group"
    $form_StartForm.Controls.Add($label_DeptGroup)

$comboBox_DeptGroup = New-ComboBox
$comboBox_DeptGroup.Size = Set-Size 240 20
$comboBox_DeptGroup.Location = Set-Point 160 220
$comboBox_DeptGroup.DropDownStyle = "DropDownList"
$comboBox_DeptGroup.Add_SelectedIndexChanged({Check-Fill})
$comboBox_DeptGroup_Choices = Import-CSV "C:\Scripts\Text Files\GoogleGroups.csv" |Where { $_.Type -eq "department" } |Sort name
ForEach ($deptGroup in $comboBox_DeptGroup_Choices) {
    $comboBox_DeptGroup.Items.Add($deptGroup.name)
}
    $form_StartForm.Controls.Add($comboBox_DeptGroup)

$label_EmployeeType = New-Label
$label_EmployeeType.Size = Set-Size 140 20
$label_EmployeeType.Location = Set-Point 10 250
$label_EmployeeType.Text = "Employee Type"
    $form_StartForm.Controls.Add($label_EmployeeType)

$comboBox_EmployeeType = New-ComboBox
$comboBox_EmployeeType.Size = Set-Size 240 20
$comboBox_EmployeeType.Location = Set-Point 160 250
$comboBox_EmployeeType.DropDownStyle = "DropDownList"
$comboBox_EmployeeType.Add_SelectedIndexChanged({Check-Fill})
$comboBox_EmployeeType_Choices =
	"Employee",
	"Intern",
	"Programista",
	"Jeavio",
	"Contractor",
	"Sub-Contractor"
ForEach ($employeeType in $comboBox_EmployeeType_Choices) {
    $comboBox_EmployeeType.Items.Add($employeeType)
}
    $form_StartForm.Controls.Add($comboBox_EmployeeType)

$label_ExpirationDate = New-Label
$label_ExpirationDate.Size = Set-Size 140 20
$label_ExpirationDate.Location = Set-Point 10 280
$label_ExpirationDate.Text = "Expiration Date"
    $form_StartForm.Controls.Add($label_ExpirationDate)

$dateTimePicker = New-DatePicker
$dateTimePicker.Size = Set-Size 240 20
$dateTimePicker.Location = Set-Point 160 280
$dateTimePicker.Enabled = $False
	$form_StartForm.Controls.Add($dateTimePicker)

$checkbox_AccountExpires = New-Checkbox
$checkbox_AccountExpires.Size = Set-Size 140 20
$checkbox_AccountExpires.Location = Set-Point 10 310
$checkbox_AccountExpires.Text = "Account Expires"
$checkbox_AccountExpires.Add_CheckStateChanged({
	If ($checkBox_AccountExpires.CheckState -eq $True) {
	$dateTimePicker.Enabled = $True
	} Else {
		$dateTimePicker.Enabled = $False
		}
	})
	$form_StartForm.Controls.Add($checkbox_AccountExpires)

$checkbox_IsManager = New-CheckBox
$checkbox_IsManager.Size = Set-Size 140 20
$checkbox_IsManager.Location = Set-Point 150 310
$checkbox_IsManager.Text = "Is Manager"
	$form_StartForm.Controls.Add($checkbox_IsManager)

$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 270 340
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$button_Cancel.Text = "Cancel"
	$form_StartForm.Controls.Add($button_Cancel)
	$form_StartForm.CancelButton = $button_Cancel

$button_OK = New-Button
$button_OK.Size = Set-Size 60 30
$button_OK.Location = Set-Point 340 340
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button_OK.Text = "OK"
$button_OK.Enabled = $False
$button_OK.Add_Click({
    $script:theUser = $comboBox_Name.SelectedItem
	$script:theManager = $textbox_Manager.Text
    $script:theTitle = $textBox_Title.Text
    $script:theDept = $comboBox_Department.SelectedItem
    $script:theOfficeLoc = $comboBox_OfficeLoc.SelectedItem
	$script:theEmployeeType = $comboBox_EmployeeType.SelectedItem
    If ($comboBox_CorpGroup.SelectedIndex -gt -1) {$script:theCorpGroup = $comboBox_CorpGroup.SelectedItem} Else {$script:theCorpGroup = "N/A"}
	If ($comboBox_LocGroup.SelectedIndex -gt -1) {$script:theLocGroup = $comboBox_LocGroup.SelectedItem} Else {$script:theLocGroup = "N/A"}
	If ($comboBox_DeptGroup.SelectedIndex -gt -1) {$script:theDeptGroup = $comboBox_DeptGroup.SelectedItem} Else {$script:theDeptGroup = "N/A"}
    If ($checkbox_IsManager.CheckState -eq "Checked") {$script:theManagerStatus = "True"}
    If ($checkBox_IsManager.CheckState -eq "Unchecked") {$script:theManagerStatus = "False"}
    If ($dateTimePicker.Enabled -eq $True) {[dateTime]$script:theExpirationDate = $dateTimePicker.Value.ToString("MM/dd/yyyy") } Else {$script:theExpirationDate = "Never"}
})
	$form_StartForm.AcceptButton = $button_OK
	$form_StartForm.Controls.Add($button_OK)

$fr = $form_StartForm.ShowDialog()
If ($fr -eq "Cancel") { Run-MainMenu }
If ($fr -eq "OK") { Confirmation-Screen }
}

	# Enable the OK button only once all fields have values

Function Check-Fill {
If (
    $comboBox_Name.SelectedIndex -gt -1 -AND
	$managerCorrect -eq "True" -AND
    $comboBox_Department.SelectedIndex -gt -1 -AND
    $comboBox_OfficeLoc.SelectedIndex -gt -1 -AND
    $comboBox_EmployeeType.SelectedIndex -gt -1 -AND
    $textBox_Title.Text) {
		$button_OK.Enabled = $True
	} Else { $button_OK.Enabled = $False }
}

	# Confirm re-provisioning

Function Confirmation-Screen {
$form_ConfirmationScreen = New-Form
$form_ConfirmationScreen.Text = "  Employment Tool"
$form_ConfirmationScreen.AutoSize = $True
$form_ConfirmationScreen.SizeGripStyle = 'Hide'
$form_ConfirmationScreen.StartPosition = 'CenterScreen'
$form_ConfirmationScreen.FormBorderStyle = 'Fixed3D'
$form_ConfirmationScreen.Font = "Arial,10,style=Bold"
$form_ConfirmationScreen.TopMost = $True
$form_ConfirmationScreen.KeyPreview	= $True
$form_ConfirmationScreen.MaximizeBox = $False
$form_ConfirmationScreen.MinimizeBox = $False
$form_ConfirmationScreen.Icon = Set-Icon

$label_theUser = New-Label
$label_theUser.AutoSize = $True
$label_theUser.Location = Set-Point 10 10
$label_theUser.Text = "Username: $theUser"
	$form_ConfirmationScreen.Controls.Add($label_theUser)

$label_theManager = New-Label
$label_theManager.AutoSize = $True
$label_theManager.Location = Set-Point 10 40
$label_theManager.Text = "Manager: $theManager"
	$form_ConfirmationScreen.Controls.Add($label_theManager)

$label_theTitle = New-Label
$label_theTitle.AutoSize = $True
$label_theTitle.Location = Set-Point 10 70
$label_theTitle.Text = "Title: $theTitle"
	$form_ConfirmationScreen.Controls.Add($label_theTitle)

$label_theDept = New-Label
$label_theDept.AutoSize = $True
$label_theDept.Location = Set-Point 10 100
$label_theDept.Text = "Department: $theDept"
	$form_ConfirmationScreen.Controls.Add($label_theDept)

$label_theOfficeLoc = New-Label
$label_theOfficeLoc.AutoSize = $True
$label_theOfficeLoc.Location = Set-Point 10 130
$label_theOfficeLoc.Text = "Office: $theOfficeLoc"
	$form_ConfirmationScreen.Controls.Add($label_theOfficeLoc)

$label_theCorpGroup = New-Label
$label_theCorpGroup.AutoSize = $True
$label_theCorpGroup.Location = Set-Point 10 160
$label_theCorpGroup.Text = "Corporate Group: $theCorpGroup"
	$form_ConfirmationScreen.Controls.Add($label_theCorpGroup)

$label_theLocGroup = New-Label
$label_theLocGroup.AutoSize = $True
$label_theLocGroup.Location = Set-Point 10 190
$label_theLocGroup.Text = "Location Group: $theLocGroup"
	$form_ConfirmationScreen.Controls.Add($label_theLocGroup)

$label_theDeptGroup = New-Label
$label_theDeptGroup.AutoSize = $True
$label_theDeptGroup.Location = Set-Point 10 220
$label_theDeptGroup.Text = "Dept Group: $theDeptGroup"
	$form_ConfirmationScreen.Controls.Add($label_theDeptGroup)

$label_theEmployeeType = New-Label
$label_theEmployeeType.AutoSize = $True
$label_theEmployeeType.Location = Set-Point 10 250
$label_theEmployeeType.Text = "Employee Type: $theEmployeeType"
	$form_ConfirmationScreen.Controls.Add($label_theEmployeeType)

$label_theManagerStatus = New-Label
$label_theManagerStatus.AutoSize = $True
$label_theManagerStatus.Location = Set-Point 10 280
$label_theManagerStatus.Text = "Is Manager? $theManagerStatus"
	$form_ConfirmationScreen.Controls.Add($label_theManagerStatus)

$label_theExpirationDate = New-Label
$label_theExpirationDate.AutoSize = $True
$label_theExpirationDate.Location = Set-Point 10 310
$label_theExpirationDate.Text = "Account Expires: $theExpirationDate"
	$form_ConfirmationScreen.Controls.Add($label_theExpirationDate)

$label_Warning = New-Label
$label_Warning.Location = Set-Point 60 350
$label_Warning.MaximumSize = New-Object System.Drawing.Size(165,40)
$label_warning.AutoSize = $True
$label_Warning.ForeColor = "Red"
$label_Warning.Font = "Arial,8,style=Italic"
$label_Warning.Text = "* By clicking OK you will begin to re-provision this user!"
	$form_ConfirmationScreen.Controls.add($label_Warning)

$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 60 30
$button_Cancel.Location = Set-Point 270 350
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_ConfirmationScreen.Controls.Add($button_Cancel)
	$form_ConfirmationScreen.CancelButton = $button_Cancel

$button_OK = New-Button
$button_OK.Size = Set-Size 60 30
$button_OK.Location = Set-Point 340 350
$button_OK.Text = "OK"
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_ConfirmationScreen.Controls.Add($button_OK)
	$form_ConfirmationScreen.AcceptButton = $button_OK

$fr = $form_ConfirmationScreen.ShowDialog()
If ($fr -eq "Cancel") { Start-Form }
If ($fr -eq "OK") { Begin-ReProvisioning }
}

	# Begin re-Provisioning

Function Begin-ReProvisioning {
	# VARs
# $theUser
# $theManager
# $theTitle
# $theDept
# $theOfficeLoc
# $theCorpGroup
# $theLocGroup
# $theDeptGroup
# $theEmployeeType
# $theManagerStatus
# $theExpirationDate
     # Configure Progress Bar Form
$form_BeginReProvisioning = New-Form
$form_BeginReProvisioning.Size = Set-Size 230 100
$form_BeginReProvisioning.Text = "  Employment Tool"
$form_BeginReProvisioning.SizeGripStyle = 'Hide'
$form_BeginReProvisioning.StartPosition = 'CenterScreen'
$form_BeginReProvisioning.FormBorderStyle = 'Fixed3D'
$form_BeginReProvisioning.Font = "Arial,10,style=Bold"
$form_BeginReProvisioning.TopMost = $True
$form_BeginReProvisioning.KeyPreview = $True
$form_BeginReProvisioning.MaximizeBox = $False
$form_BeginReProvisioning.MinimizeBox = $False
$form_BeginReProvisioning.Icon = Set-Icon

$label = New-Label
$label.Size = Set-Size 200 20
$label.Location = Set-Point 10 10
$label.Text = "Preparing to re-provision..."
	$form_BeginReProvisioning.Controls.Add($label)

$progression = New-ProgressBar
$progression.Size = Set-Size 200 20
$progression.Location = Set-Point 10 40
$progression.Value = 0
	$form_BeginReProvisioning.Controls.Add($progression)

$form_BeginReProvisioning.Show()
Sleep 2
    # Set Confirm Preference to None so that no CLI prompts appear
$ConfirmPreference 	= 'None'
    # Enable the account
$label.Text = "Enabling AD Account..."
$progression.Value = 5
$form_BeginReProvisioning.Refresh()
Set-ADUser $theUser -Enabled $True
    # Move to Users OU
$label.Text = "Moving to Users OU..."
$progression.Value = 10
$form_BeginReProvisioning.Refresh()
Get-ADUser $theUser |Move-ADObject -TargetPath "CN=Users,DC=sevone,DC=com"
    # Set Day1@SevOne! password
$label.Text = "Setting password..."
$progression.Value = 15
$form_BeginReProvisioning.Refresh()
$pw = "Day1@SevOne!"
$spw = ConvertTo-SecureString $pw -AsPlainText -Force
Set-ADAccountPassword $theUser -NewPassword $spw
    # Remove/Set Expiration
If ($theExpirationDate -eq "Never") {
    $label.Text = "Clearing Expiration..."
    $progression.Value = 20
    $form_BeginReProvisioning.Refresh()
    Clear-ADAccountExpiration $theUser
} Else {
    $label.Text = "Setting Expiration..."
    $progression.Value = 20
    $form_BeginReProvisioning.Refresh()
    Set-ADAccountExpiration $theUser $theExpirationDate.AddDays(1)
}
    # Set Title
$label.Text = "Setting Title..."
$progression.Value = 25
$form_BeginReProvisioning.Refresh()
Set-ADUser $theUser -Title $theTitle
    # Set Department
$label.Text = "Setting Department..."
$progression.Value = 30
$form_BeginReProvisioning.Refresh()
Set-ADUser $theUser -Department $theDept
	# Set Description
$label.Text = "Setting Description..."
$progression.Value = 35
$form_BeginReProvisioning.Refresh()
$theDescription = "$theEmployeeType - $theDept"
Set-ADUser $theUser -Description $theDescription
	# Set Office
$label.Text = "Setting Office..."
$progression.Value = 40
$form_BeginReProvisioning.Refresh()
Set-ADUser $theUser -Office $theOfficeLoc
	# Set Manager
$label.Text = "Setting Manager..."
$progression.Value = 43
$form_BeginReProvisioning.Refresh()
Set-ADUser $theUser -Manager $theManager
	# Re-enable home drive archive
$testTheArchive = Test-Path "\\fileserver.sevone.com\File Server\home\Former Employee Archives\$theUser"
$testTheCurrent = Test-Path "\\fileserver.sevone.com\File Server\home\$theUser"
$theUserInfo = Get-ADUser $theUser -Property *
If ($testTheArchive -eq $True -AND $testTheCurrent -eq $False) {
	$label.Text = "Unarchiving Home Folder..."
	$progression.Value = 45
	$form_BeginReProvisioning.Refresh()
	Move-Item -Path "\\fileserver.sevone.com\File Server\home\Former Employee Archives\$theUser" "\\fileserver.sevone.com\File Server\home\"
}
If ($testTheArchive -eq $False -AND $testTheCurrent -eq $False) {
	$label.Text = "Creating Home Folder..."
	$progression.Value = 45
	$form_BeginReProvisioning.Refresh()
	New-Item -Path "\\fileserver.sevone.com\File Server\home\$theUser" -Type Directory
}
If ($testTheCurrent -eq $True) {
	[System.Windows.Forms.MessageBox]::SHOW("This user has a home drive folder that was never archived or conflicts with another username:`n`n       \\fileserver.sevone.com\File Server\home\$theUser" , "Warning" , "OK" , "Warning")
}
If ($testTheCurrent -eq $False -AND $userInfo.HomeDirectory -ne "\\fileserver.sevone.com\File Server\home\$theUser") {
	Set-ADUser $theUser -HomeDrive "Z:" -HomeDirectory "\\fileserver.sevone.com\File Server\home\$theUser"
	$theHomeACL = Get-ACL "\\fileserver.sevone.com\File Server\home\$theUser"
	$theNewAR = New-Object System.Security.AccessControl.FileSystemAccessRule("$theUser","FullControl","Allow")
	$theHomeACL.SetAccessRule($theNewAR)
	Set-ACL "\\fileserver.sevone.com\File Server\home\$theUser" $theHomeACL
}
	# Set standard AD groups
$label.Text = "Modifying AD groups..."
$progression.Value = 60
$form_BeginReProvisioning.Refresh()
Add-ADGroupMember "SevOne" $theUser
Add-ADGroupMember "NonAdmins" $theUser
Add-ADGroupMember "SSLVPN" $theUser
Add-ADGroupMember "Wireless" $theUser
Add-ADGroupMember "Domain Users" $theUser
	# Set employment type specific groups
If ($theEmployeeType -eq "Employee") {Add-ADGroupMember "SevOne Staff" $theUser} # This group enables access to SevOne Source for SevOne Full-Time Employees (FTEs)
If ($theEmployeeType -eq "Programista") {Add-ADGroupMember "Programista" $theUser} # This group disables access to SevOne Source for all Programista contractors
If ($theEmployeeType -eq "Jeavio") {Add-ADGroupMember "Jeavio" $theUser} # This group disables access to SevOne Source for all Jeavio contractors
If ($theEmployeeType -eq "Contractor" -OR $theEmployeeType -eq "Sub-Contractor") {Add-ADGroupMember "Contractors" $theUser} # Just a contractor group, nothing special
	# Change primary group membership to Domain Users
$theGroupToken = Get-ADGroup "Domain Users" -Properties @("primaryGroupToken")
Get-ADUser $theUser |Set-ADUser -Replace @{primaryGroupID=$theGroupToken.primaryGroupToken}
	# Remove from Former Employees AD Group
Remove-ADGroupMember "Former Employees" $theUser
	# Move to 2FA Exception OU in Google
$label.Text = "Google GAM stuff..."
$progression.Value = 75
$form_BeginReProvisioning.Refresh()
GAM update user $theUser ou "/005 - 2FA Exception"
	# Add to Google Groups (Corporate, Location, Department)
If ($theCorpGroup -eq "N/A") {} Else {GAM update group $theCorpGroup add user $theUser}
If ($theDeptGroup -eq "N/A") {} Else {GAM update group $theDeptGroup add user $theUser}
If ($theLocGroup -ne "N/A (International)" -OR $theLocGroup -ne "N/A") {GAM update group $theLocGroup add user $theUser}
	# Remove vacation responder
GAM user $theUser vacation off
    # Remove Delegates
$theDelegates = GAM user $theUser show delegates |Select-String " Delegate Email: " | %{ $_ -creplace ' Delegate Email: '}
ForEach ($delegate in $theDelegates) {
     GAM user $theUser delete delegate $delegate
}
     # Google Profile shared
GAM user $theUser profile shared
# Method to turn off 2 step verification enrollment???
# None :(

# Enable samanage account
	$label.Text = "Re-Enabling Samanage..."
	$progression.Value = 90
	$form_BeginReProvisioning.Refresh()
$headers = @{
     'Accept' = 'application/XML'
     'Content-Type' = 'text/XML'
}
$key = (3,4,2,3,56,34,254,222,1,1,2,23,42,54,33,233,1,34,2,7,6,5,35,43)
$credPW = cat "C:\Scripts\Text Files\APIcreds.txt" |convertto-securestring -key $key
$creds = New-Object -typename System.Management.Automation.PSCredential -argumentlist "sevoneadmin@sevone.com",$credPW
$GetURI = "https://api.samanage.com/users.xml?email=$theUser@sevone.com"

# GET the user - simply used to gather the userID for the PUT request
[string]$XMLResults 	=(Invoke-WebRequest $GetURI -credential $creds -method get)
$XMLResults |Out-File C:\Temp\XMLResults.xml
$id = Get-Content C:\Temp\XMLResults.xml |Select-String ID |Select -First 1 | % { $_ -creplace '[^0-9]' }
Remove-Item C:\Temp\XMLResults.xml

# PUT the change 
$PutURI = "https://api.samanage.com/users/$id.xml"
[string]$reprovXML =
    "<user>
     <id>$id</id>
     <name>$theUser@sevone.com</name>
     <disabled>false</disabled>
    </user>"

Invoke-RestMethod -Body $reprovXML -Credential $creds -Headers $headers -Method PUT -uri $PutURI

# Clean Home Brew files
$label.Text = "Cleaning homebrew..."
$progression.Value = 95
$form_BeginReProvisioning.Refresh()
$theFormerArchive = Get-Content 'C:\Scripts\Text Files\FormerEmployeeUsernamesArchive.txt' |Where { $_ -notmatch "$theUser" }
$theFormerArchive |Out-File 'C:\Scripts\Text Files\FormerEmployeeUsernamesArchive.txt' -Force
$theCurrentStaff = Get-Content 'C:\Scripts\Text Files\All Employees List.txt'
$theCurrentStaff + $theUser |Sort |Out-File "C:\Scripts\Text Files\All Employees List.txt"
	# Log use
	$label.Text = "Logging use..."
	$progression.Value	= 98
	$form_BeginReProvisioning.Refresh()

	
## WebEx API code to check if user had an account. 
## If so, re-enable.
## If not, but needs one, create.
	
$date	= Get-Date
Echo "$date :: $loggedinUser re-provisioned the account for $theUser." >> $logFile

$label.Text = "Re-provisioning Complete!"
$progression.Value = 100
$form_BeginReProvisioning.Refresh()
Sleep 5
$form_BeginReProvisioning.Close()
Run-MainMenu
}

Start-Form
