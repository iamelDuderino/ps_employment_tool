# Employment Tool v3.0
# 	Google Groups - Create

	# Gather group full name and short name

Function Start-Form {

$groupsList = GAM print groups |Select -Skip 1 | % { $_ -creplace '@sevone.com' }
$usersList = GAM print users |Select -Skip 1 | % { $_ -creplace '@sevone.com' }
$checkList = $groupsList + $usersList

$form_StartForm = New-Form
$form_StartForm.Size = Set-Size 390 170
$form_StartForm.Text = "  Employment Tool"
$form_StartForm.SizeGripStyle = 'Hide'
$form_StartForm.StartPosition = 'CenterScreen'
$form_StartForm.FormBorderStyle = 'Fixed3D'
$form_StartForm.Font = "Arial,10,style=Bold"
$form_StartForm.TopMost = $True
$form_StartForm.KeyPreview = $True
$form_StartForm.MaximizeBox = $False
$form_StartForm.MinimizeBox = $False
$form_StartForm.Icon = Set-Icon

$label_theFullName = New-Label
$label_theFullName.Size = Set-Size 150 20
$label_theFullName.Location	= Set-Point 10 10
$label_theFullName.Text = "Group Display Name"
	$form_StartForm.Controls.Add($label_theFullName)

$textBox_theFullName = New-TextBox
$textBox_theFullName.Size = Set-Size 200 20
$textBox_theFullName.Location = Set-Point  170 10
$textBox_theFullName.Add_TextChanged({Check-Fill})
	$form_StartForm.Controls.Add($textBox_theFullName)

$label_theShortName = New-Label
$label_theShortName.Size = Set-Size 150 20
$label_theShortName.Location = Set-Point 10 40
$label_theShortName.Text = "Group Short Name"
	$form_StartForm.Controls.Add($label_theShortName)

$textBox_theShortName = New-TextBox
$textBox_theShortName.Size = Set-Size 200 20
$textBox_theShortName.Location = Set-Point 170 40
$textBox_theShortName.Add_TextChanged({
	If ($textBox_theShortName.Text -ne "") {
		$script:groupCheck = $checkList |Select-String "^($($textBox_theShortName.Text))\b" -Quiet
		If ($groupCheck -eq $True) {
			$label_theShortName.Forecolor = "Red" # Can not create because it already exists
			Check-Fill
			}
		If ($groupCheck -ne $True) {
			$label_theShortName.Forecolor = "Green" # Can proceed with creation because group name does not exist
			Check-Fill
			}
		} Else {
			$label_theShortName.Forecolor = "Black" # Set back to default
			Check-Fill
			}
	})
	$form_StartForm.Controls.Add($textBox_theShortName)

$checkbox_PublicPosting = New-Checkbox
$checkbox_PublicPosting.Size = Set-Size 150 20
$checkbox_PublicPosting.Location = Set-Point 250 70
$checkbox_PublicPosting.Text = "Public Posting"
	$form_StartForm.Controls.Add($checkbox_PublicPosting)

$button_Create = New-Button
$button_Create.Size = Set-Size 70 30
$button_Create.Location = Set-Point 300 100
$button_Create.Text = "Create"
$button_Create.Enabled = $False
$button_Create.DialogResult	= "OK"
$button_Create.Add_Click({
	$script:groupFullName = $textBox_theFullName.Text
	$script:groupShortName = $textBox_theShortName.Text
	If ($checkbox_PublicPosting.Checked -eq $True) {$script:groupPublicPosting = $True}
	If ($checkbox_PublicPosting.Checked -eq $False) {$script:groupPublicPosting = $False}
	})
	$form_StartForm.AcceptButton = $button_Create
	$form_StartForm.Controls.Add($button_Create)

$button_Cancel = New-Button
$button_Cancel.Size = Set-Size 70 30
$button_Cancel.Location = Set-Point 220 100
$button_Cancel.Text = "Cancel"
$button_Cancel.DialogResult	= "Cancel"
	$form_StartForm.CancelButton = $button_Cancel
	$form_StartForm.Controls.Add($button_Cancel)


$fr = $form_StartForm.ShowDialog()
If ($fr -eq "Cancel") { $form_StartForm.Dispose() ; Run-MainMenu }
If ($fr -eq "OK") { $form_StartForm.Dispose() ; Create-Group }
}

	# Enable OK once all fields are filled out

Function Check-Fill {
	If ($textBox_theShortName.Text -ne "" -AND $textBox_theFullName.Text -ne "" -AND $groupCheck -ne $True) {
		$button_Create.Enabled = $True
	} ElseIf ($textBox_theShortName.Text -eq "" -OR $textBox_theFullName.Text -eq "" -OR $groupCheck -eq $True) {
		$button_Create.Enabled = $False
	}
}

	# Create the Group

Function Create-Group {
GAM Create Group $groupShortName Max_Message_Bytes 25M Is_Archived False Name "$($groupFullName)" Allow_External_Members False Message_Moderation_Level Moderate_None Show_In_Group_Directory True Who_Can_Invite All_Managers_Can_Invite Who_Can_Join Invited_Can_Join Who_Can_View_Group All_Members_Can_View Allow_Web_Posting False Members_Can_Post_As_The_Group False
If ($groupPublicPosting -eq $True) {gam update group $groupShortName Who_Can_Post_Message Anyone_Can_Post}
If ($groupPublicPosting -eq $False) {gam update group $groupShortName Who_Can_Post_Message All_In_Domain_Can_Post}
Write-Log "$loggedInUser created the group $($groupFullName) ($($groupShortName))"
$addMembers = [System.Windows.Forms.MessageBox]::SHOW("$($groupFullName) ($($groupShortName)) created!`n`nClick OK to add members to new group or Cancel to abort. " , "Success" , "OKCancel")
If ($addMembers -eq "OK") {Add-Members}
If ($addMembers -eq "Cancel") {Run-MainMenu}
}

	# Add members to new group

Function Add-Members {
$form_AddMembers = New-Form
$form_AddMembers.Size = Set-Size 400 280
$form_AddMembers.Text = "  Employment Tool"
$form_AddMembers.SizeGripStyle = 'Hide'
$form_AddMembers.StartPosition = 'CenterScreen'
$form_AddMembers.FormBorderStyle = 'Fixed3D'
$form_AddMembers.Font = "Arial,10,style=Bold"
$form_AddMembers.TopMost = $True
$form_AddMembers.KeyPreview = $True
$form_AddMembers.MaximizeBox = $False
$form_AddMembers.MinimizeBox = $False
$form_AddMembers.Icon = Set-Icon

$label_AddUsers = New-Label
$label_AddUsers.Size = Set-Size 100 20
$label_AddUsers.Location = Set-Point 10 20
$label_AddUsers.Text = "Add user(s)"
	$form_AddMembers.Controls.Add($label_AddUsers)

$label_Warning = New-Label
$label_Warning.Size = Set-Size 250 20
$label_Warning.Location = Set-Point 170 180
$label_Warning.ForeColor = "Red"
$label_Warning.Text = "* Separate multiple users with a comma."
$label_Warning.Font = "Arial,8,style=Italic"
	$form_AddMembers.Controls.Add($label_Warning)

$textbox_AddUsers = New-RichTextBox
$textbox_AddUsers.Size = Set-Size 250 150
$textbox_AddUsers.Location	= Set-Point 120 20
$textbox_AddUsers.Add_TextChanged({
	If ($textbox_AddUsers.Text -ne "") {
		$button_OK.Enabled = $True
		} ElseIf ($textbox_AddUsers.Text -eq "") {
			$button_OK.Enabled = $False
			}
	})
	$form_AddMembers.Controls.Add($textbox_AddUsers)

$progressbar_AddUsers = New-ProgressBar
$progressbar_AddUsers.Size	= Set-Size 200 20
$progressbar_AddUsers.Location	= Set-Point 10 220
$progressbar_AddUsers.Value	= 0
	$form_AddMembers.Controls.Add($progressbar_AddUsers)

$button_OK = New-Button
$button_OK.Size = Set-Size 50 30
$button_OK.Location = Set-Point 320 210
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button_OK.Text = "OK"
$button_OK.Enabled = $False
$button_OK.Add_Click({
	$button_OK.Enabled = $False
	$usersToAdd = $textbox_AddUsers.Text -creplace "\s" |% { $_.Split(",") }
	$userList = gam print users |Select -Skip 1
	$progressbar_AddUsers.Value	= 25
	$groupShortNameList	= gam print groups |Select -Skip 1
	$progressbar_AddUsers.Value	= 50
	# Increase value for progress bar
	$increaseValue	= 50/$usersToAdd.Count
	$n		= 0
	ForEach ($user in $usersToAdd) {
		$checkUsers	= $userList |Select-String $user -Quiet
		$checkGroups	= $groupShortNameList |Select-String $user -Quiet
		If (!$checkUsers -AND !$checkGroups) {
			$notFoundMsg = Throw-Error "The user ($user) was not found in the system.`n`nClick OK to add anyway or CANCEL to discard this addition."
			If ($notFoundMsg -eq "OK") {
				gam update group $groupShortName add user $user
				Write-Log "$loggedInUser added the user $user to $groupShortName."
				}
			If ($notFoundMsg -eq "Cancel") {}
			} Else {
				gam update group $groupShortName add user $user
				$n++
				$progressbar_AddUsers.Value = $increaseValue*$n +50
				$date = Get-Date
				Write-Log "$loggedInUser added the user $user to $groupShortName."
				}
	}
	Sleep 2
	$script:newGroupMembers = gam info group $groupShortName |findstr " member: " |%{ $_.Replace(" member: ","") } |%{ $_ -creplace '@sevone.*' }
	})
	$form_AddMembers.Controls.Add($button_OK)
	$form_AddMembers.AcceptButton = $button_OK

$fr = $form_AddMembers.ShowDialog()
If ($fr -eq "Cancel") {$form_AddMembers.Dispose() ; Run-MainMenu}
If ($fr -eq "OK") {$form_AddMembers.Dispose() ; Show-Results}
}

	# After member additions, show membership

Function Show-Results {
$form_ShowResults = New-Form
$form_ShowResults.Size = Set-Size 400 260
$form_ShowResults.Text = "  Employment Tool"
$form_ShowResults.SizeGripStyle	= 'Hide'
$form_ShowResults.StartPosition	= 'CenterScreen'
$form_ShowResults.FormBorderStyle = 'Fixed3D'
$form_ShowResults.Font = "Arial,10,style=Bold"
$form_ShowResults.TopMost = $True
$form_ShowResults.KeyPreview = $True
$form_ShowResults.MaximizeBox = $False
$form_ShowResults.MinimizeBox = $False
$form_ShowResults.Icon = Set-Icon

$label_Results = New-Label
$label_Results.Size = Set-Size 100 20
$label_Results.Location = Set-Point 10 20
$label_Results.Text = "Member List:"
	$form_ShowResults.Controls.Add($label_Results)

$listbox_Results = New-ListBox
$listbox_Results.Size = Set-Size 250 150
$listbox_Results.Location = Set-Point 120 20
$listbox_Results.SelectionMode = 'None'
ForEach ($member in $newGroupMembers) {
	$listbox_Results.Items.Add($member)
	}
	$form_ShowResults.Controls.Add($listbox_Results)

$button_OK = New-Button
$button_OK.Size = Set-Size 50 30
$button_OK.Location = Set-Point 320 190
$button_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button_OK.Text = "OK"
$button_OK.Enabled = $True
	$form_ShowResults.Controls.Add($button_OK)

$fr = $form_ShowResults.ShowDialog()
If ($fr -eq "OK") {$form_ShowResults.Dispose() ; Run-MainMenu}
}

Start-Form