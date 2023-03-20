# Employment Tool v3.0
# 	Employee Accounts - Employee Changes

	# Gather information for employment information modifications

Function Start-Form {

$ADList = Get-ADUser -Filter * -SearchBase "CN=Users,DC=domain,DC=com" |Select SamAccountName

$form_StartForm = New-Form
$form_StartForm.Text = "  Employment Tool - Changes"
$form_StartForm.Size = Set-Size 480 410
$form_StartForm.SizeGripStyle = 'Hide'
$form_StartForm.StartPosition = 'CenterScreen'
$form_StartForm.FormBorderStyle	= 'Fixed3D'
$form_StartForm.Font = "Arial,10,style=Bold"
$form_StartForm.TopMost = $True
$form_StartForm.KeyPreview = $True
$form_StartForm.MaximizeBox = $False
$form_StartForm.MinimizeBox = $False
$form_StartForm.Icon = Set-Icon

$label_Username = New-Label
$label_Username.Size = Set-Size 130 20
$label_Username.Location = Set-Point 10 10
$label_Username.Text = "Username"
	$form_StartForm.Controls.Add($label_Username)

$textbox_Username = New-Textbox
$textbox_Username.Size = Set-Size 300 20
$textbox_Username.Location = Set-Point 150 10
$textbox_Username.Add_TextChanged({
	If ($textbox_Username.Text -ne "") {
		$checkAD = $ADList.SamAccountName |Select-String "^$($textbox_Username.Text)\b$" -Quiet
		If ($checkAD -eq $True) {
			$label_Username.ForeColor = "Green"
			Fill-Form
			} Else {
			$label_Username.ForeColor = "Red"
			Clear-Form
			}
		}
	})
	$form_StartForm.Controls.Add($textbox_Username)

$label_Title = New-Label
$label_Title.Size = Set-Size 130 20
$label_Title.Location = Set-Point 10 40
$label_Title.Text = "Title"
	$form_StartForm.Controls.Add($label_Title)

$textbox_Title = New-Textbox
$textbox_Title.Size = Set-Size 300 20
$textbox_Title.Location = Set-Point 150 40
$textbox_Title.Add_TextChanged({
	$script:theNewTitle = $textbox_Title.Text
})
	$form_StartForm.Controls.Add($textbox_Title)

$label_Manager = New-Label
$label_Manager.Size = Set-Size 130 20
$label_Manager.Location = Set-Point 10 70
$label_Manager.Text = "Manager"
	$form_StartForm.Controls.Add($label_Manager)

$textbox_Manager = New-Textbox
$textbox_Manager.Size = Set-Size 300 20
$textbox_Manager.Location = Set-Point 150 70
$textbox_Manager.Add_TextChanged({
	$checkAD = $ADList.SamAccountName |Select-String "^$($textbox_Manager.Text)\b$" -Quiet
	If ($checkAD -eq $True) {
		$label_Manager.ForeColor = "Green"
		$script:theNewManager = $textbox_Manager.Text
		$button_Submit.Enabled = $True
		} Else {
		$label_Manager.ForeColor = "Red"
		$script:theNewManager = $theCurrentManager
		$button_Submit.Enabled = $False
		}
	})
	$form_StartForm.Controls.Add($textbox_Manager)

$label_Dept = New-Label
$label_Dept.Size = Set-Size 130 20
$label_Dept.Location = Set-Point 10 100
$label_Dept.Text = "Department"
	$form_StartForm.Controls.Add($label_Dept)

$combobox_Dept = New-Combobox
$combobox_Dept.Size = Set-Size 300 20
$combobox_Dept.Location = Set-Point 150 100
$combobox_Dept.DropDownStyle = "DropDownList"
$combobox_Dept.Add_SelectedIndexChanged({
	$script:theNewDept = $combobox_Dept.SelectedItem
	})
$deptList = Get-Content "C:\Scripts\Text Files\DeptCodes.txt"
ForEach ($dept in $deptList) {
	$combobox_Dept.Items.Add($dept)
	}
	$form_StartForm.Controls.Add($combobox_Dept)

$label_OfficeLoc = New-Label
$label_OfficeLoc.Size = Set-Size 130 20
$label_OfficeLoc.Location = Set-Point 10 130
$label_OfficeLoc.Text = "Office Location"
	$form_StartForm.Controls.Add($label_OfficeLoc)

$combobox_OfficeLoc = New-Combobox
$combobox_OfficeLoc.Size = Set-Size 300 20
$combobox_OfficeLoc.Location = Set-Point 150 130
$combobox_OfficeLoc.DropDownStyle = "DropDownList"
$combobox_OfficeLoc.Add_SelectedIndexChanged({
	$script:theNewOfficeLoc = $combobox_OfficeLoc.SelectedItem
	})
$officeList = Get-Content "C:\Scripts\Text Files\Offices.txt"
ForEach ($office in $officeList) {
		$combobox_OfficeLoc.Items.Add($office)
	}
	$form_StartForm.Controls.Add($combobox_OfficeLoc)

$label_CorpGroup = New-Label
$label_CorpGroup.Size = Set-Size 130 20
$label_CorpGroup.Location = Set-Point 10 160
$label_CorpGroup.Text = "Corporate Group"
	$form_StartForm.Controls.Add($label_CorpGroup)

$combobox_CorpGroup = New-Combobox
$combobox_CorpGroup.Size = Set-Size 300 20
$combobox_CorpGroup.Location = Set-Point 150 160
$combobox_CorpGroup.DropDownStyle = "DropDownList"
$combobox_CorpGroup.Add_SelectedIndexChanged({
	$script:theNewCorpGroup = $combobox_CorpGroup.SelectedItem
	})
$corpGroupList = Import-CSV "C:\Scripts\Text Files\GoogleGroups.csv" |Where {$_.Type -match "corporate"}
ForEach ($corpGroup in $corpGroupList) {
	$combobox_CorpGroup.Items.Add($corpGroup.Name)
	}
	$form_StartForm.Controls.Add($combobox_CorpGroup)

$label_LocGroup = New-Label
$label_LocGroup.Size = Set-Size 130 20
$label_LocGroup.Location = Set-Point 10 190
$label_LocGroup.Text = "Location Group"
	$form_StartForm.Controls.Add($label_LocGroup)

$combobox_LocGroup = New-Combobox
$combobox_LocGroup.Size = Set-Size 300 20
$combobox_LocGroup.Location = Set-Point 150 190
$combobox_LocGroup.DropDownStyle = "DropDownList"
$combobox_LocGroup.Add_SelectedIndexChanged({
	$script:theNewLocGroup = $combobox_LocGroup.SelectedItem
	})
$locGroupList = Import-CSV "C:\Scripts\Text Files\GoogleGroups.csv" |where {$_.Type -match "location"} |Sort name
ForEach ($locGroup in $locGroupList) {
		$combobox_LocGroup.Items.Add($locGroup.Name)
	}
	$form_StartForm.Controls.Add($combobox_LocGroup)

$label_DeptGroup = New-Label
$label_DeptGroup.Size = Set-Size 130 20
$label_DeptGroup.Location = Set-Point 10 220
$label_DeptGroup.Text = "Department Group"
	$form_StartForm.Controls.Add($label_DeptGroup)

$combobox_DeptGroup = New-Combobox
$combobox_DeptGroup.Size = Set-Size 300 20
$combobox_DeptGroup.Location = Set-Point 150 220
$combobox_DeptGroup.DropDownStyle = "DropDownList"
$combobox_DeptGroup.Add_SelectedIndexChanged({
	$script:theNewDeptGroup = $combobox_DeptGroup.SelectedItem
})
$deptGroupList = Import-CSV "C:\Scripts\Text Files\GoogleGroups.csv" |Where {$_.Type -match "department"} |Sort name
ForEach ($deptGroup in $deptGroupList) {
	$combobox_DeptGroup.Items.Add($deptGroup.Name)
	}
	$form_StartForm.Controls.Add($combobox_DeptGroup)


$label_EmployeeType = New-Label
$label_EmployeeType.Size = Set-Size 130 20
$label_EmployeeType.Location = Set-Point 10 250
$label_EmployeeType.Text = "Employee Type"
	$form_StartForm.Controls.Add($label_EmployeeType)	

$combobox_EmployeeType = New-Combobox
$combobox_EmployeeType.Size = Set-Size 300 20
$combobox_EmployeeType.Location = Set-Point 150 250
$combobox_EmployeeType.DropDownStyle = "DropDownList"
$combobox_EmployeeType.Add_SelectedIndexChanged({
	$script:theNewEmployeeType = $combobox_EmployeeType.SelectedItem
	})
$employeeTypeList =
    "Employee",
    "Programista",
	"Jeavio",
    "Contractor",
    "Sub-Contractor",
    "Intern"
ForEach ($type in $employeeTypeList) {$combobox_EmployeeType.Items.Add($type)}
	$form_StartForm.Controls.Add($combobox_EmployeeType)

$checkbox_AccountExpires = New-Checkbox
$checkbox_AccountExpires.Size = Set-Size 140 20
$checkbox_AccountExpires.Location = Set-Point 10 310
$checkbox_AccountExpires.Text = "Account Expires"
$checkbox_AccountExpires.Add_CheckStateChanged({
	If ($checkbox_AccountExpires.CheckState -eq $True) {
		$dateTime_Expiration.Enabled = $True
		} Else {
			$dateTime_Expiration.Enabled = $False
			}
	})
	$form_StartForm.Controls.Add($checkbox_AccountExpires)
	
$checkbox_ManagerStatus = New-Checkbox
$checkbox_ManagerStatus.Size = Set-Size 140 20
$checkbox_ManagerStatus.Location = Set-Point 150 310
$checkbox_ManagerStatus.Text = "Is Manager"
$checkbox_ManagerStatus.Add_CheckedChanged({
	If ($checkbox_ManagerStatus.Checked -eq $True) {$script:theNewManagerStatus = $True}
	If ($checkbox_ManagerStatus.Checked-eq $False) {$script:theNewManagerStatus = $False}	
	})
	$form_StartForm.Controls.Add($checkbox_ManagerStatus)

$label_Expiration = New-Label
$label_Expiration.Size = Set-Size 130 20
$label_Expiration.Location = Set-Point 10 280
$label_Expiration.Text = "Expiration Date"
	$form_StartForm.Controls.Add($label_Expiration)

$dateTime_Expiration = New-DatePicker
$dateTime_Expiration.Size = Set-Size 300 20
$dateTime_Expiration.Location = Set-Point 150 280
$dateTime_Expiration.Enabled = $False
$dateTime_Expiration.Add_ValueChanged({
	$script:theNewExpirationDate = $dateTime_Expiration.Value
	})
	$form_StartForm.Controls.Add($dateTime_Expiration)

$button_Quit = New-Button
$button_Quit.Size = Set-Size 70 30
$button_Quit.Location = Set-Point 300 340
$button_Quit.Text = "Quit"
	$form_StartForm.Controls.Add($button_Quit)
	$form_StartForm.CancelButton = $button_Quit

$button_Submit = New-Button
$button_Submit.Size = Set-Size 70 30
$button_Submit.Location = Set-Point 380 340
$button_Submit.Text = "Submit"
$button_Submit.Enabled = $False
$button_Submit.DialogResult = [System.Windows.Forms.DialogResult]::OK
	$form_StartForm.Controls.Add($button_Submit)
	$form_StartForm.AcceptButton = $button_Submit


$fr = $form_StartForm.ShowDialog()
If ($fr -eq "Cancel") { $form_StartForm.Dispose() ; Run-Synchronization }
If ($fr -eq "OK") { $form_StartForm.Dispose() ; Update-User }
}

	# Fills out form with current information

Function Fill-Form {
$button_Submit.Enabled = $False
$button_Quit.Enabled = $False
$textbox_Username.Enabled = $False
$textbox_Title.Enabled = $False
$textbox_Manager.Enabled = $False
$combobox_Dept.Enabled = $False
$combobox_OfficeLoc.Enabled = $false
$combobox_CorpGroup.Enabled = $False
$combobox_LocGroup.Enabled = $False
$combobox_DeptGroup.Enabled = $False
$combobox_EmployeeType.Enabled = $False

Sleep 3

		$script:theUser = $textbox_Username.Text
		$theUserInformation = Get-ADUser $theUser -Properties *
		$theManagerInformation = Get-ADUser $theUserInformation.Manager
			$textbox_Title.Text = $theUserInformation.Title
			$textbox_Manager.Text = $theManagerInformation.SamAccountName
			$combobox_Dept.SelectedItem = $theUserInformation.Department
			$combobox_OfficeLoc.SelectedItem = $theUserInformation.Office
			$combobox_EmployeeType.SelectedItem = $theUserInformation.Description -creplace '.-.*'
				$script:theCurrentTitle = $theUserInformation.Title
				$script:theCurrentManager = $theManagerInformation.SamAccountName
				$script:theCurrentDept = $theUserInformation.Department
				$script:theCurrentOfficeLoc = $theUserInformation.Office				
				$script:theCurrentEmployeeType = $theUserInformation.Description -creplace '.-.*'
				$script:theCurrentExpirationDate = $theUserInformation.AccountExpirationDate
$theUsersGroups	= gam info user $theUser |findstr "@domain.com>" | % { $_ -creplace '.*<' } | % { $_ -creplace '@domain.com>' }
ForEach ($locGroup in $locGroupList.Name) {
	$check = $theUsersGroups |Select-String $locGroup\b -Quiet
	If ($check) {
		$combobox_LocGroup.SelectedItem = $locGroup
		$script:theCurrentLocGroup = $locGroup
		}
	}

ForEach ($corpGroup in $corpGroupList.Name) {
	$check = $theUsersGroups |Select-String $corpGroup\b -Quiet
	If ($check) {
		$combobox_CorpGroup.SelectedItem = $corpGroup
		$script:theCurrentCorpGroup = $corpGroup
		}
	}

ForEach ($deptGroup in $deptGroupList.Name) {
	$check = $theUsersGroups |Select-String $deptGroup\b -Quiet
	If ($check) {
		$combobox_DeptGroup.SelectedItem = $deptGroup
		$script:theCurrentDeptGroup = $deptGroup
		}
	}

$managerList = ((gam info group Managers |Select-String " member: ") -creplace ' member: ') -creplace ' \(user\)'
$checkManagers = $managerList |Select-String $theUser\b -Quiet
If ($checkManagers -eq $True) {
	$checkbox_ManagerStatus.Checked = $True
	$script:theCurrentManagerStatus	= $True
	}
If (!$checkManagers) {
	$checkbox_ManagerStatus.Checked = $False
	$script:theCurrentManagerStatus	= $False
	$script:theNewManagerStatus		= $False
	}
If ($theUserInformation.AccountExpirationDate) {
	$checkbox_AccountExpires.Checked = $True
	$dateTime_Expiration.Enabled = $True
	$dateTime_Expiration.Value = $theCurrentExpirationDate
	} Else {
		$checkbox_AccountExpires.Checked = $False
		$dateTime_Expiration.Enabled = $False
		}
$button_Submit.Enabled = $True
$button_Quit.Enabled = $True
$textbox_Username.Enabled = $True
$textbox_Title.Enabled = $True
$textbox_Manager.Enabled = $True
$combobox_Dept.Enabled = $True
$combobox_OfficeLoc.Enabled = $True
$combobox_CorpGroup.Enabled = $True
$combobox_LocGroup.Enabled = $True
$combobox_DeptGroup.Enabled = $True
$combobox_EmployeeType.Enabled = $True
$textbox_Title.Focus()
}

	# Clear the form if the user is not found

Function Clear-Form {
$textbox_Title.Clear()
$textbox_Manager.Clear()
$combobox_Dept.SelectedIndex = -1
$combobox_OfficeLoc.SelectedIndex = -1
$combobox_CorpGroup.SelectedIndex = -1
$combobox_LocGroup.SelectedIndex = -1
$combobox_DeptGroup.SelectedIndex = -1
$combobox_EmployeeType.SelectedIndex = -1
$checkbox_ManagerStatus.Checked = $False
$checkbox_AccountExpires.Checked = $False
$date = Get-Date
$dateTime_Expiration.Value = $date
}

	# Update user with modified information

Function Update-User {
$form_UpdateUser = New-Form
$form_UpdateUser.Size = Set-Size 290 140
$form_UpdateUser.Text = "  Employment Tool - Employee Changes"
$form_UpdateUser.SizeGripStyle = 'Hide'
$form_UpdateUser.StartPosition = 'CenterScreen'
$form_UpdateUser.FormBorderStyle = 'Fixed3D'
$form_UpdateUser.Font = "Arial,10,style=Bold"
$form_UpdateUser.TopMost = $True
$form_UpdateUser.KeyPreview = $True
$form_UpdateUser.MaximizeBox = $False
$form_UpdateUser.MinimizeBox = $False
$form_UpdateUser.Icon = Set-Icon

$label_Status = New-Label
$label_Status.Size = Set-Size 250 20
$label_Status.Location = Set-Point 10 10
$label_Status.Text = "Updating user..."
	$form_UpdateUser.Controls.Add($label_Status)

$progressBar_UpdateUser = New-ProgressBar
$progressBar_UpdateUser.Size = Set-Size 250 20
$progressBar_UpdateUser.Location = Set-Point 10 40
$progressBar_UpdateUser.Value = 0
	$form_UpdateUser.Controls.Add($progressBar_UpdateUser)
	
$button_OK = New-Button
$button_OK.Size = Set-Size 50 20
$button_OK.Location = Set-Point 210 80
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button_OK.Text = "OK"
	$form_UpdateUser.Controls.Add($button_OK)
	$form_UpdateUser.AcceptButton	= $button_OK

	# Set new descr with selected employee type / dept
If ($theNewEmployeeType -AND $theNewDept) {
	$theNewDescr = "$theNewEmployeeType - $theNewDept"
	Set-ADUser $theUser -Description "$theNewDescr"
	}
If ($theNewEmployeeType -AND !$theNewDept) {
	$theNewDescr = "$theNewEmployeeType - $theCurrentDept"
	Set-ADUser $theUser -Description "$theNewDescr"
	}
If (!$theNewEmployeeType -AND $theNewDept) {
	$theNewDescr = "$theCurrentEmployeeType - $theNewDept"
	Set-ADUser $theUser -Description "$theNewDescr"
	}

	# Set new title, if new title

If ($theCurrentTitle -ne $theNewTitle) {
	Set-ADUser $theUser -Title $theNewTitle
	$date = Get-Date
	Echo "$date :: $loggedInUser modified $theUser's title to $theNewTitle" >> $logFile
	}
$label_Status.Text				= "Title checked/updated..."
$progressBar_UpdateUser.Value	= 10
$form_UpdateUser.Refresh()

	# Set new manager, if new manager

If ($theCurrentManager -ne $theNewManager) {
	Set-ADUser $theUser -Manager $theNewManager
	$date = Get-Date
	Echo "$date :: $loggedInUser modified $theUser's manager to $theNewManager" >> $logFile
	}
$label_Status.Text				= "Manager checked/updated..."
$progressBar_UpdateUser.Value	= 20
$form_UpdateUser.Refresh()

	# Set new dept, if new dept

If ($theCurrentDept -ne $theNewDept) {
	Set-ADUser $theUser -Department $theNewDept
	GAM update user $theUser org "/$theNewDept"
	$date = Get-Date
	Echo "$date :: $loggedInUser modified $theUser's dept to $theNewDept" >> $logFile
	}
$label_Status.Text				= "Dept checked/updated..."
$progressBar_UpdateUser.Value	= 30
$form_UpdateUser.Refresh()

	# Set new office loc, if new office loc

If ($theCurrentOfficeLoc -ne $theNewOfficeLoc) {
	Set-ADUser $theUser -Office $theNewOfficeLoc
	$date = Get-Date
	Echo "$date :: $loggedInUser modified $theUser's office location to $theNewOfficeLoc" >> $logFile
	}
$label_Status.Text = "Office Location checked/updated..."
$progressBar_UpdateUser.Value = 40
$form_UpdateUser.Refresh()

	# Set new corp group, if new corp group

If ($theCurrentCorpGroup -ne $theNewCorpGroup) {
	gam update group $theCurrentCorpGroup remove user $theUser
	gam update group $theNewCorpGroup add user $theUser
	$date = Get-Date
	Echo "$date :: $loggedInUser modified $theUser's corp group to $theNewCorpGroup" >> $logFile
	}
$label_Status.Text = "Corp Group checked/updated..."
$progressBar_UpdateUser.Value = 50
$form_UpdateUser.Refresh()

	# Set new loc group, if new loc group

If ($theCurrentLocGroup -ne $theNewLocGroup -AND $theNewLocGroup -ne "N/A International") {
	gam update group $theCurrentLocGroup remove user $theUser
	gam update group $theNewLocGroup add user $theUser
	$date = Get-Date
	Echo "$date :: $loggedInUser modified $theUser's loc group to $theNewLocGroup" >> $logFile
	}
$label_Status.Text = "Loc Group checked/updated..."
$progressBar_UpdateUser.Value = 60
$form_UpdateUser.Refresh()

	# Set new dept group, if new dept group

If ($theCurrentDeptGroup -ne $theNewDeptGroup) {
	gam update group $theCurrentDeptGroup remove user $theUser
	gam update group $theNewDeptGroup add user $theUser
	$date = Get-Date
	Echo "$date :: $loggedInUser modified $theUser's dept group to $theNewDeptGroup" >> $logFile
	}
$label_Status.Text = "Dept Group checked/updated..."
$progressBar_UpdateUser.Value = 70
$form_UpdateUser.Refresh()

# Set new employee type, if new employee type
If ($theCurrentEmployeeType -ne $theNewEmployeeType) {
	If ($theNewEmployeeType -eq "Employee") {
		$currentADGroups = Get-ADPrincipalGroupMembership "$theUser" |Select Name
		If ($currentADGroups -notcontains "domain Staff") {Add-ADGroupMember "domain Staff" $theUser}
		If ($currentADGroups -contains "Programista") {Remove-ADGroupMember "Programista" $theUser}
		If ($currentADGroups -contains "Contractors") {Remove-ADGroupMember "Contractors" $theUser}
		}
	If ($theNewEmployeeType -eq "Contractor" -OR $theNewEmployeeType -eq "Sub-Contractor") {
		$currentADGroups = Get-ADPrincipalGroupMembership "$theUser" |Select Name
		If ($currentADGroups -notcontains "Contractors") {Add-ADGroupMember "Contractors" $theUser}
		If ($currentADGroups -contains "Programista") {Remove-ADGroupMember "Programista" $theUser}
		}
	If ($theNewEmployeeType -eq "Intern") {
		$currentADGroups = Get-ADPrincipalGroupMembership "$theUSer" |Select Name
		If ($currentADGroups -notcontains "domain Staff") {Add-ADGroupMember "domain Staff" $theUser}
		If ($currentADGroups -contains "Programista") {Remove-ADGroupMember "Programista" $theUser}
		If ($currentADGroups -contains "Contractors") {Remove-ADGroupMember "Contractors" $theUser}
		}
	If ($theNewEmployeeType -eq "Programista") {
		$currentADGroups = Get-ADPrincipalGroupMembership "$theUser" |Select Name
		If ($currentADGroups -notcontains "Programista") {Add-ADGroupMember "Programista" $theUser}
		If ($currentADGroups -contains "Contractors") {Remove-ADGroupMember "Contractors" $theUser}
		If ($currentADGroups -contains "domain Staff") {Remove-ADGroupMember "domain Staff" $theUSer}
		}
	If ($theNewEmployeeType -eq "Jeavio") {
		$currentADGroups = Get-ADPrincipalGroupMembership "$theUser" |Select Name
		If ($theCurrentADGroups -notcontains "Jeavio") {Add-ADGroupMember "Jeavio" $theUser}
		If ($theCurrentADGroups -contains "Contractors") {Remove-ADGroupMember "Contractors" $theUser}
		If ($theCurrentADGroups -contains "domain Staff") {Remove-ADGroupMember "domain Staff" $theUser} 
		}
	$date = Get-Date
	Echo "$date :: $loggedInUser modified $theUser's employee type to $theNewEmployeeType"
}

	# Set new expiration date, if new expiration date

If ($theCurrentExpirationDate -ne $theNewExpirationDate -AND $dateTime_Expiration.Enabled -eq $True) {
	Set-ADAccountExpiration $theUser $theNewExpirationDate.AddDays(1)
	}
If ($dateTime_Expiration.Enabled -eq $False) {
	Clear-ADAccountExpiration $theUSer
	}

$label_Status.Text				= "Employee Type checked/updated..."
$progressBar_UpdateUser.Value	= 80
$form_UpdateUser.Refresh()

	# Set new manager status, if new manager status

If ($theCurrentManagerStatus -ne $theNewManagerStatus) {
	If ($theNewManagerStatus -eq $False) {
		gam update group managers remove user $theUser
		$date = Get-Date
		Echo "$date :: $loggedInUser modified $theUser's manager status to False and removed from group managers@domain.com" >> $logFile
		}
	If ($theNewManagerStatus -eq $True) {
		gam update group managers add user $theUser
		$date = Get-Date
		Echo "$date :: $loggedInUser modified $theUser's manager status to True and added to group managers@domain.com" >> $logFile
		}
	}
$label_Status.Text = "Manager status checked/updated..."
$progressBar_UpdateUser.Value = 90
$form_UpdateUser.Refresh()

$label_Status.Text =  "User $theUser updated!"
$progressBar_UpdateUser.Value = 100
$form_UpdateUser.Refresh()

$fr = $form_UpdateUser.ShowDialog()
If ($fr -eq "OK") { $form_UpdateUser.Dispose() ; Start-Form }
If ($fr -eq "Cancel") {$form_UpdateUser.Dispose() }
}

	# Synchronize all of the changes to Google

Function Run-Synchronization {
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
$label_GSync.Text = "Syncing to google."
	$form_GSync.Controls.Add($label_GSync)

$form_GSync.Show()
Sync-Google
Sleep 1
Do {
	$form_GSync.Refresh()
	$label_GSync.Text = "Syncing to google."
	Sleep 1
	$form_GSync.Refresh()
	$label_GSync.Text = "Syncing to google.."
	Sleep 1
	$form_GSync.Refresh()
	$label_GSync.Text = "Syncing to google..."
	Sleep 1
	$form_GSync.Refresh()
	$label_GSync.Text = "Syncing to google...."
	Sleep 1
	$form_GSync.Refresh()
	$checkProcess = Get-Process |Select ProcessName |Where {$_.ProcessName -match "sync-cmd"}
	} Until (!$checkProcess)

$label_GSync.Text = "Done!"
$form_GSync.Refresh()
Sleep 2
$form_GSync.Dispose()
Run-MainMenu
}

Start-Form