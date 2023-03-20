# Employment Tool v3.0
# 	Google Groups - Add/Remove/Modifications

	# Query group

Function Start-Form {
$form_StartForm	= New-Form
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

	# Which group?

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Size = Set-Size 200 30
$textBox1.Location = Set-Point 110 10
$textBox1.Add_TextChanged({
	If ($textBox1.Text -ne "") {$queryBtn.Enabled = $True} Else {$queryBtn.Enabled = $False}
	})
	$form_StartForm.Controls.Add($textBox1)

$label1 = New-Label
$label1.Size = Set-Size 100 30
$label1.Location = Set-Point 10 10
$label1.Text = "Which group?"
	$form_StartForm.Controls.Add($label1)

$queryBtn = New-Button
$queryBtn.Size = Set-Size 100 23
$queryBtn.Location = Set-Point 210 40
$queryBtn.Text = "Run Query"
$queryBtn.Enabled = $False
$queryBtn.Add_Click({
	$queryResults.Items.Clear()
	$script:group = $textBox1.Text
	$progressBar1.Value	= 30
	$queriedGroup = gam info group $group 
	$exitCode = $?
	$progressBar1.Value	= 60
	If ($exitCode -eq $False) {
		[System.Windows.Forms.MessageBox]::Show("$group was not found." , "Error" , "OK")
		$removeUser.Enabled = $False
		$addUser.Enabled	= $False
		$progressBar1.Value = 0
		}
	If ($exitCode -eq $True) {
	If ($group -match "^delt*$") {
		[System.Windows.Forms.MessageBox]::SHOW("DELT memberships must be approved by DELT members!" , "Warning" , "OK")
		}
	$progressBar1.Value	= 80
	$members = $queriedGroup |findstr " member: " |%{ $_.Replace(" member: ","") } |%{ $_ -creplace '@sevone.*' }
	ForEach ($member in $members) {
		$queryResults.Items.Add($member)
		}
	$progressBar1.Value = 100
	$addUser.Enabled	= $True
	$publicPostingCheck = $queriedGroup |Select-String "whoCanPostMessage" |% { $_ -creplace ".*:." }
	If ($publicPostingCheck -eq "ANYONE_CAN_POST") {
		$checkbox_PublicPosting.CheckState = "Checked"
		} Else {
			$checkbox_PublicPosting.CheckState = "Unchecked"
			}
	$progressBar1.Value = 0
		}
	})
	$form_StartForm.Controls.Add($queryBtn)
	$form_StartForm.AcceptButton	= $queryBtn

$script:queryResults	= New-ListBox
$queryResults.Size	= Set-Size 230 225
$queryResults.Location	= Set-Point 80 70
$queryResults.SelectionMode	= "MultiExtended"
$queryResults.Add_SelectedIndexChanged({
	$removeUser.Enabled	= $True
	})
	$form_StartForm.Controls.Add($queryResults)

$checkbox_PublicPosting = New-Checkbox
$checkbox_PublicPosting.Size = Set-Size 150 20
$checkbox_PublicPosting.Location = Set-Point 190 290
$checkbox_PublicPosting.Text = "Public Posting"
$checkbox_PublicPosting.Add_CheckStateChanged({
	If ($checkbox_PublicPosting.CheckState -eq "Checked") {
		gam update group $group who_can_post_message anyone_can_post
		}
	If ($checkbox_PublicPosting.CheckState -eq "Unchecked") {
		gam update group $group who_can_post_message all_in_domain_can_post
		}
	})
	$form_StartForm.Controls.Add($checkbox_PublicPosting)

$script:progressBar1 = New-ProgressBar
$progressBar1.Value	= 0
$progressBar1.Size	= Set-Size 230 20
$progressBar1.Location	= Set-Point 80 320 	
	$form_StartForm.Controls.Add($progressBar1)

$addUser = New-Button
$addUser.Size = Set-Size 70 30
$addUser.Location = Set-Point 240 350
$addUser.Text = "Add"
$addUser.Enabled = $False
$addUser.Add_Click({
	Add-User
	})
	$form_StartForm.Controls.Add($addUser)

$removeUser	= New-Button
$removeUser.Size = Set-Size 70 30
$removeUser.Location = Set-Point 160 350
$removeUser.Text = "Remove"
$removeUser.Enabled	= $False
$removeUser.Add_Click({
	Remove-User
	})
	$form_StartForm.Controls.Add($removeUser)

$cancel = New-Button
$cancel.Size = Set-Size 70 30
$cancel.Location = Set-Point 80 350
$cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$cancel.Text = "Cancel"
	$form_StartForm.Controls.Add($cancel)
	$form_StartForm.CancelButton = $cancel

$fr = $form_StartForm.ShowDialog()
If ($fr -eq "Cancel") { $form_StartForm.Dispose() ; Run-MainMenu }
}

	# Add user(s)

Function Add-User {
$form_AddUser = New-Form
$form_AddUser.Size = Set-Size 400 280
$form_AddUser.Text = "  Employment Tool"
$form_AddUser.SizeGripStyle	= 'Hide'
$form_AddUser.StartPosition	= 'CenterScreen'
$form_AddUser.FormBorderStyle = 'Fixed3D'
$form_AddUser.Font = "Arial,10,style=Bold"
$form_AddUser.TopMost = $True
$form_AddUser.KeyPreview = $True
$form_AddUser.MaximizeBox = $False
$form_AddUser.MinimizeBox = $False
$form_AddUser.Icon = Set-Icon

$label1 = New-Label
$label1.Size = Set-Size 100 20
$label1.Location = Set-Point 10 20
$label1.Text = "Enter user(s)"
	$form_AddUser.Controls.Add($label1)

$label2 = New-Label
$label2.Size = Set-Size 250 20
$label2.Location = Set-Point 170 180
$label2.Text = "* Separate multiple users with a comma."
$label2.Font = "Arial,8,style=Italic"
	$form_AddUser.Controls.Add($label2)

$textBox1 = New-RichTextBox
$textBox1.Size = Set-Size 250 150
$textBox1.Location = Set-Point 120 20
$textBox1.Add_TextChanged({
	If ($textBox1.Text -ne "") {
		$ok.Enabled = $True
		} ElseIf ($textBox1.Text -eq "") {
		$ok.Enabled = $False
		}
	})
	$form_AddUser.Controls.Add($textBox1)
	
$progressBar2 = New-ProgressBar
$progressBar2.Size	= Set-Size 200 20
$progressBar2.Location	= Set-Point 10 220
$progressBar2.Value	= 0
	$form_AddUser.Controls.Add($progressBar2)

$ok = New-Button
$ok.Size = Set-Size 50 30
$ok.Location = Set-Point 320 210
$ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
$ok.Text = "OK"
$ok.Enabled = $False
$ok.Add_Click({
	$ok.Enabled = $False
	$usersToAdd	= $textBox1.Text -creplace "\s" |% { $_.Split(",") }
	$userList	= gam print users |Select -Skip 1
	$progressBar2.Value	= 25
	$groupList	= gam print groups |Select -Skip 1
	$progressBar2.Value	= 50
	# Increase value for progress bar
	$increaseValue	= 50/$usersToAdd.Count
	$n		= 0
	ForEach ($user in $usersToAdd) {
		$checkUsers	= $userList |Select-String $user -Quiet
		$checkGroups	= $groupList |Select-String $user -Quiet
		If (!$checkUsers -AND !$checkGroups) {
			$notFoundMsg = Throw-Error "The user ($user) was not found in the system.`n`nClick OK to add anyway or CANCEL to discard this addition."
			If ($notFoundMsg -eq "OK") {
				gam update group $group add user $user
				Write-Log "$loggedInUser added the user $user to $group."
				}
			If ($notFoundMsg -eq "Cancel") {}
			} Else {
				gam update group $group add user $user
				$n++
				$progressBar2.Value = $increaseValue*$n +50
				$date = Get-Date
				Write-Log "$loggedInUser added the user $user to $group."
			}
		}
	$queryResults.Items.Clear()
	$members = gam info group $group |findstr " member: " |%{ $_.Replace(" member: ","") } |%{ $_ -creplace '@sevone.*' }
	ForEach ($member in $members) {
		$queryResults.Items.Add($member)
		}
	$form_AddUser.Dispose()
	})
	$form_AddUser.Controls.Add($ok)
	$form_AddUser.AcceptButton = $ok

$fr = $form_AddUser.ShowDialog()
If ($fr -eq "Cancel") { $form_AddUser.Dispose() }
}

	# Remove user(s)

Function Remove-User {
$usersToRemove	= $queryResults.SelectedItems
$n 		= 0
$increaseValue 	= 100/$usersToRemove.Count
ForEach ($user in $queryResults.SelectedItems) {
	$n++
	$progressBar1.Value	= $increaseValue*$n
	gam update group $group remove user $user
	$date	= Get-Date
	Echo "$date :: $loggedinUser removed the user $user from $group" >> $logFile
	}
$queryResults.Items.Clear()
$members = gam info group $group |findstr " member: " |%{ $_.Replace(" member: ","") } |%{ $_ -creplace '@sevone.*' }
ForEach ($member in $members) {
	$queryResults.Items.Add($member)
	}
$progressBar1.Value = 0
}

Start-Form
