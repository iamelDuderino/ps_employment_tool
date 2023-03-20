# Employment Tool v3.0
# 	Main Menu
#
# This tool is meant to serve as a GUI integration for Active Directory & Google Environments for the SevOne Help Desk team.
#
# Google GAM maintained @ GitHub by jay0lee
# 	https://github.com/jay0lee/GAM
#
# Google GCDS (formerly known as GADS) maintained @ Google by Google
#	https://support.google.com/a/answer/106368?hl=en
#
# Employment Tool maintained @ IT-GADS by AJ Tomko (until termination lulz!)
#	https://plus.google.com/+AJTomko

Import-Module EmploymentTool_v3.0

# Function for below: Allow access if signed in user is allowed to RDP and has GADS configured
Function Main-Menu {

## BUTTONS/LABELS
$form_MainMenu = New-Form
$form_MainMenu.Size = Set-Size 480 310
$form_MainMenu.Text = "  Employment Tool - Main Menu"
$form_MainMenu.SizeGripStyle = 'Hide'
$form_MainMenu.StartPosition = 'CenterScreen'
$form_MainMenu.FormBorderStyle	= 'Fixed3D'
$form_MainMenu.Font = "Arial,10,style=Bold"
$form_MainMenu.TopMost = $True
$form_MainMenu.KeyPreview = $True
$form_MainMenu.MaximizeBox = $False
$form_MainMenu.MinimizeBox = $False
$form_MainMenu.Icon = Set-Icon

$label_ChooseCategory = New-Label
$label_ChooseCategory.Text = "Choose A Category:"
$label_ChooseCategory.AutoSize = $True
$label_ChooseCategory.Location = Set-Point 10 10
	$form_MainMenu.Controls.Add($label_ChooseCategory)

$listbox_Category = New-ListBox
$listbox_Category.Size = Set-Size 200 200
$listbox_Category.Location = Set-Point 30 30
$listbox_CategoryOptions =
	"Employee Accounts", # cr 0
	"Service Accounts", #cr 1
	"Google Groups", # cr 2
	"Tools" # cr 3
	#"Inventory" # cr 4
ForEach ($choice in $listbox_CategoryOptions) {
	$listbox_Category.Items.Add($choice)
}
	$form_MainMenu.Controls.Add($listbox_Category)

$listbox_SubCategory = New-ListBox
$listbox_SubCategory.TabIndex = 0
$listbox_SubCategory.Size = Set-Size 200 200
$listbox_SubCategory.Location = Set-Point 250 30
	$form_MainMenu.Controls.Add($listbox_SubCategory)

$button_cancel = New-Button
$button_cancel.TabIndex = 1
$button_cancel.Text = "Cancel"
$button_cancel.Size = Set-Size 60 30
$button_cancel.Location = Set-Point 325 240
$button_cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_MainMenu.Controls.Add($button_cancel)
	$form_MainMenu.CancelButton	= $button_cancel

$button_ok = New-Button
$button_ok.TabIndex = 2
$button_ok.Text = "OK"
$button_ok.Size = Set-Size 60 30
$button_ok.Location = Set-Point 390 240
$button_ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button_ok.Enabled = $False
	$form_MainMenu.Controls.Add($button_ok)
	$form_MainMenu.AcceptButton	= $button_ok
	
## CATEGORY EVENTS

$listbox_Category.Add_SelectedIndexChanged({

	# If category = Employee Accounts (cr0)

If ($listbox_Category.SelectedIndex -eq 0) {
	$listbox_SubCategory.Items.Clear()
	$listbox_SubCategoryOptions_EmployeeAccounts =
		"Provision User", # CR0SCR0
		"Deprovision User", # CR0SCR1
		"Reprovision User", # CR0SCR2
		"Employee Changes" # CR0SCR3
	ForEach ($choice in $listbox_SubCategoryOptions_EmployeeAccounts) {
		$listbox_SubCategory.Items.Add($choice)
		}
	}

	# If category = Service Accounts (cr1)

If ($listbox_Category.SelectedIndex -eq 1) {
	$listbox_SubCategory.Items.Clear()
	$listbox_SubCategoryOptions_ServiceAccounts =
		"AD - Create", # CR1SCR0
		"AD - Delete", # CR1SCR1
		"Google - Create", # CR1SCR2
		"Google - Delete", # CR1SCR3
		"Google - Add Delegate", # CR1SCR4
		"Google - Remove Delegate" # CR1SCR5
	ForEach ($choice in $listbox_SubCategoryOptions_ServiceAccounts) {
		$listbox_SubCategory.Items.Add($choice)
		}
	}

	# If category = Google Groups (cr2)

If ($listbox_Category.SelectedIndex -eq 2) {
	$listbox_SubCategory.Items.Clear()
	$listbox_SubCategoryOptions_GoogleGroups =
		"Create", # CR2SCR0
		"Delete", # CR2SCR1
		"Add/Remove/Modify" # CR2SCR2
	ForEach ($choice in $listbox_SubCategoryOptions_GoogleGroups) {
		$listbox_SubCategory.Items.Add($choice)
		}
	}

	# If category = Tools (cr3)

If ($listbox_Category.SelectedIndex -eq 3) {
	$listbox_SubCategory.Items.Clear()
	$listbox_SubCategoryOptions_Tools =
		"Create JIRA Account", # CR3SCR0
		"Create WebEx Account", # CR3SCR1
		"Update Phone Info", # CR3SCR2
		"Add User To Druva InSync", # CR3SCR3
		"Add User To Encryption", #CR3SCR4
		"Add User To GIT/SRC" #CR3SCR5
	ForEach ($choice in $listbox_SubCategoryOptions_Tools) {
		$listbox_SubCategory.Items.Add($choice)
		}
	}

If ($listbox_Category.SelectedIndex -eq 4) {
	$listbox_SubCategory.Items.Clear()
	$listbox_SubCategoryOptions_Inventory = 
		"Newark", # CR4SCR0
		"Boston" # CR4SCR1
	ForEach ($choice in $listbox_SubCategoryOptions_Inventory) {
		$listbox_SubCategory.Items.Add($choice)
		}
	}

	# Enable ok button once a sub category has been selected

$listbox_SubCategory.Add_SelectedIndexChanged({
	If ($listbox_SubCategory.SelectedIndex -gt -1) {$button_ok.Enabled = $True} Else {$button_ok.Enabled = $False}
	})
})

	## RUN FORM

$fr = $form_MainMenu.ShowDialog()

	## RESULTS

If ($fr -eq "Cancel"){ $form_MainMenu.Dispose() }
If ($fr -eq "OK") {
$cr = $listbox_Category.SelectedIndex
$scr = $listbox_SubCategory.SelectedIndex

	## DONE!! 
	
If ($cr -eq 0 -AND $scr -eq 0) {$form_MainMenu.Dispose() ; Run-EAccounts-Provision}
If ($cr -eq 0 -AND $scr -eq 1) {$form_MainMenu.Dispose() ; Run-EAccounts-Deprovision}
If ($cr -eq 0 -AND $scr -eq 2) {$form_MainMenu.Dispose() ; Run-EAccounts-Reprovision}
If ($cr -eq 0 -AND $scr -eq 3) {$form_MainMenu.Dispose() ; Run-EAccounts-Changes}

If ($cr -eq 1 -AND $scr -eq 0) {$form_MainMenu.Dispose() ; }
If ($cr -eq 1 -AND $scr -eq 1) {$form_MainMenu.Dispose() ; }
If ($cr -eq 1 -AND $scr -eq 2) {$form_MainMenu.Dispose() ; Run-SAccounts-G-Create}
If ($cr -eq 1 -AND $scr -eq 3) {$form_MainMenu.Dispose() ; Run-SAccounts-G-Delete}
If ($cr -eq 1 -AND $scr -eq 4) {$form_MainMenu.Dispose() ; Run-SAccounts-G-AddDelegate}
If ($cr -eq 1 -AND $scr -eq 5) {$form_MainMenu.Dispose() ; Run-SAccounts-G-RemoveDelegate}

If ($cr -eq 2 -AND $scr -eq 0) {$form_MainMenu.Dispose() ; Run-GoogleGroups-Create}
If ($cr -eq 2 -AND $scr -eq 1) {$form_MainMenu.Dispose() ; Run-GoogleGroups-Delete}
If ($cr -eq 2 -AND $scr -eq 2) {$form_MainMenu.Dispose() ; Run-GoogleGroups-AddRemoveModify}

If ($cr -eq 3 -AND $scr -eq 0) {$form_MainMenu.Dispose() ; Run-CreateJIRAAccount}
If ($cr -eq 3 -AND $scr -eq 1) {$form_MainMenu.Dispose() ; Run-CreateWebExAccount}
If ($cr -eq 3 -AND $scr -eq 2) {$form_MainMenu.Dispose() ; Run-PhoneUpdater}
If ($cr -eq 3 -AND $scr -eq 3) {$form_MainMenu.Dispose() ; Run-AddToDruva}
If ($cr -eq 3 -AND $scr -eq 4) {$form_MainMenu.Dispose() ; Run-AddToEncryption}
If ($cr -eq 3 -AND $scr -eq 5) {$form_MainMenu.Dispose() ; Run-AddToSRC}

If ($cr -eq 4 -AND $scr -eq 0) {$form_MainMenu.Dispose() ; Run-Inventory-Newark}
If ($cr -eq 4 -AND $scr -eq 1) {$form_MainMenu.Dispose() ; }

	## INCOMPLETE!!

If ($cr -eq 1 -AND $scr -eq 0) { Throw-Error "This form is not ready yet." ; $form_MainMenu.Dipose() ; Run-MainMenu }
If ($cr -eq 1 -AND $scr -eq 1) { Throw-Error "This form is not ready yet." ; $form_MainMenu.Dipose() ; Run-MainMenu }

# Hidden Tools Menu

	## END!!
	}
}

	# Function for below: Unauthorized access if user is not allowed to RDP to the system
	
Function Unauthorized-Access {
Throw-Error "You are not authorized to use this application!"
}

	# Function for below: Deny access if the user is in the RDP list but has yet to configure GADS (mandatory)
	
Function Deny-Access {
Throw-Error "You have not set up your GADS profile!`n`nPlease review the format here:`n`nhttps://sites.google.com/a/sevone.com/internal-it/account-management/it-gads---profile-configuration"
}

	# Function for below: Deny access if the user is not in the Domain Admins group

Function Not-DomainAdmin {
Throw-Error "Your account is not in the Domain Admins group. The tool may malfunction depending on permissions. Aborting."
}

	# Pre-check

$GADSCheck			= Test-Path "C:\GADS\Account-and-Profiles-Sync-$loggedInUser"
$authList 			= Net LocalGroup "Remote Desktop Users" |Where { $_ -match "SEVONE" }
$domainAdminList	= Get-ADGroupMember "Domain Admins" | %{ Get-ADUser $_ |Select SamAccountName }
$domainAdminCheck	= $domainAdminList |Select-String $loggedInUser -Quiet

	# Allow access if signed in user is allowed to RDP and has GADS configured

If ($loggedInUser -match $authList.Members -AND $GADSCheck -eq $True -AND $domainAdminCheck -eq $True) {
	Main-Menu
}

	# Deny access if not Domain Admin. Working to clean up a proper Help Desk Admins role.

If ($domainAdminCheck -eq $False) {
	Not-DomainAdmin
}

	# Unauthorized access if user is not allowed to RDP to the system

If ($loggedInUser -notmatch $authList.Members) {
	Unauthorized-Access
}

	# Deny access if the user is in the RDP list but has yet to configure GADS (mandatory)

If ($loggedInUser -match $authList.Members -AND $GADSCheck -eq $False) {
	Deny-Access
}