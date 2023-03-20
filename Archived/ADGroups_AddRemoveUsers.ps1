# Employment Tool v3
#	AD Groups - Add/Remove Users

	# Function to start form. Query group, add/remove user

$script:adGroupsList = Get-ADGroup -Filter *
Function Start-Form {
$form_StartForm = New-Form
$form_StartForm.Size = Set-Size 340 430
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


$label_GroupName = New-Label
$label_GroupName.Size = Set-Size 100 30
$label_GroupName.Location = Set-Point 10 10
$label_GroupName.Text = "Which group?"
	$form_StartForm.Controls.Add($label_GroupName)

$textbox_GroupName = New-TextBox
$textbox_GroupName.Size = Set-Size 200 30
$textbox_GroupName.Location = Set-Point 110 10
$textbox_GroupName.Add_TextChanged({
	If ($textbox_GroupName -ne "") {$button_Query.Enabled = $True} Else {$button_Query.Enabled = $False}
})
	$form_StartForm.Controls.Add($textbox_GroupName)

$button_Query = New-Button
$button_Query.Size = Set-Size 100 23
$button_Query.Location = Set-Point 210 40
$button_Query.Text = "Run Query"
$button_Query.Enabled = $False
$button_Query.Add_Click({
	$listbox_QueryResults.Items.Clear()
	$script:group = $textbox_GroupName.Text
	$checkADGroups = $adGroupsList.Name |Select-String $group -Quiet
	If ($checkADGroups -eq $True) {
		$groupMembers = Get-ADGroupMember $group | % { Get-ADUser $_.distinguishedName |Select SamAccountName }
		ForEach ($member in $groupMembers) {
			$listbox_QueryResults.Items.Add($member.SamAccountName)
		}
	} Else { Throw-Errow "The group $($textbox_GroupName.Text) was not found!" }
})
	$form_StartForm.Controls.ADd($button_Query)
	
$listbox_QueryResults = New-ListBox
$listbox_QueryResults.Size = Set-Size 230 225
$listbox_QueryResults.Location = Set-Point 80 70
$listbox_QueryResults.SelectionMode = "MultiExtended"
$listbox_QueryResults.Add_SelectedIndexChanged({
	$button_Remove.Enabled = $True
})
	$form_StartForm.Controls.Add($listbox_QueryResults)

$fr = $form_StartForm.ShowDialog()
If ($fr -eq "Cancel") { $form_StartForm.Dispose() }
}

Start-Form