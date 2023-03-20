# Employment Tool v3.0
# 	Google Groups - Delete Group

	# Gather information for which group to delete

Function Start-Form {
$form_StartForm = New-Object System.Windows.Forms.Form
$form_StartForm.Text = "  Employment Tool"
$form_StartForm.Size = Set-Size 470 130
$form_StartForm.SizeGripStyle = 'Hide'
$form_StartForm.StartPosition = 'CenterScreen'
$form_StartForm.FormBorderStyle = 'Fixed3D'
$form_StartForm.Font = "Arial,10,style=Bold"
$form_StartForm.TopMost	= $True
$form_StartForm.KeyPreview	= $True
$form_StartForm.MaximizeBox	= $False
$form_StartForm.MinimizeBox	= $False
$form_StartForm.Icon = Set-Icon

$label_groupQuery = New-Label
$label_groupQuery.Size	= Set-Size 140 20
$label_groupQuery.Location	= Set-Point 10 10
$label_groupQuery.Text	= "Which Group?"
	$form_StartForm.Controls.Add($label_groupQuery)

$dropdown_groupQuery_List = GAM print groups |Select -Skip 1
$dropdown_groupQuery = New-ComboBox
$dropdown_groupQuery.Size = Set-Size 300 20
$dropdown_groupQuery.Location = Set-Point 150 10
$dropdown_groupQuery.DropDownStyle = "DropDownList"
ForEach ($group in $dropdown_groupQuery_List) {
	$dropdown_groupQuery.Items.Add($group)
	}
	$form_StartForm.Controls.Add($dropdown_groupQuery)

$ok	= New-Button
$ok.Size = Set-Size 60 30
$ok.Location = Set-Point 390 50
$ok.Text = "OK"
$ok.DialogResult =[System.Windows.Forms.DialogResult]::OK
$ok.Add_Click({
	$script:groupToDelete = $dropdown_groupQuery.SelectedItem
	})
	$form_StartForm.Controls.Add($ok)

$cancel = New-Button
$cancel.Size = Set-Size 60 30
$cancel.Location = Set-Point 320 50
$cancel.Text = "Cancel"
$cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
	$form_StartForm.Controls.Add($cancel)

$fr = $form_StartForm.ShowDialog()
If ($fr -eq "Cancel") { Run-MainMenu }
If ($fr -eq "OK") { Delete-Group }
}

	# Delete the group

Function Delete-Group {
$deletionConfirmation = [System.Windows.Forms.MessageBox]::SHOW("Are you sure you wish to DELETE $($groupToDelete)?" , "Warning" , "OKCancel" , "Warning")
If ($deletionConfirmation -eq "Cancel") { Start-Form }
If ($deletionConfirmation -eq "OK") { GAM delete group $groupToDelete ; Start-Form }
}

Start-Form