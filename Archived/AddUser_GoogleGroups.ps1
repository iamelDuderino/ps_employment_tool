## Google Groups - Add User

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Set-Alias C:\GAM\gam.exe gam

Function Start-Form {
$form					= New-Object System.Windows.Forms.Form
$form.Size				= New-Object System.Drawing.Size(375,210)
$form.Icon				= New-Object System.Drawing.Icon("C:\Scripts\Images\domain\domainLogo1.ico")
$form.Text				= "  Employment Tool"
$form.SizeGripStyle		= 'Hide'
$form.StartPosition		= 'CenterScreen'
$form.FormBorderStyle	= 'Fixed3D'
$form.Font				= "Arial,10,style=Bold"
$form.TopMost			= $True
$form.KeyPreview		= $True
$form.MaximizeBox		= $False
$form.MinimizeBox		= $False

$label					= New-Object System.Windows.Forms.Label
$label.Size				= New-Object System.Drawing.Size(200,20)
$label.Location			= New-Object System.Drawing.Point(10,10)
$label.Text				= "Google Groups - Add User "
	$form.Controls.Add($label)

$textBox1 				= New-Object System.Windows.Forms.TextBox
$textBox1.Size 			= New-Object System.Drawing.Size(200,20)
$textBox1.Location 		= New-Object System.Drawing.Point(150,40)
$textBox1.TabIndex		= 0
	$form.Controls.Add($textBox1)
	
$textBox1Label			= New-Object System.Windows.Forms.Label
$textBox1Label.Size		= New-Object System.Drawing.Size(200,20)
$textBox1Label.Location = New-Object System.Drawing.Point(50,40)
$textBox1Label.Text		= "Which Group?"
	$form.Controls.Add($textBox1Label)
	
$textBox2				= New-Object System.Windows.Forms.TextBox
$textBox2.Size			= New-Object System.Drawing.Size(200,20)
$textBox2.Location		= New-Object System.Drawing.Point(150,70)
$textBox2.TabIndex		= 1
	$form.Controls.Add($textBox2)

$textBox2Label			= New-Object System.Windows.Forms.Label
$textBox2Label.Size		= New-Object System.Drawing.Size(200,20)
$textBox2Label.Location	= New-Object System.Drawing.Point(50,70)
$textBox2Label.Text		= "Which User(s)?"
	$form.Controls.Add($textBox2Label)

$note					= New-Object System.Windows.Forms.Label
$note.Size				= New-Object System.Drawing.Size(300,30)
$note.Location			= New-Object System.Drawing.Point(150,110)
$note.Text				= "* Separate multiple users with a comma."
$note.Font				= "Arial,8,style=Italic"
	$form.Controls.Add($note)

$cancel					= New-Object System.Windows.Forms.Button
$cancel.Size			= New-Object System.Drawing.Size(60,30)
$cancel.Location		= New-Object System.Drawing.Point(225,140)
$cancel.Text			= "Cancel"
$cancel.TabIndex 		= 2
$cancel.DialogResult	= [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($cancel)
	
$ok						= New-Object System.Windows.Forms.button
$ok.Size				= New-Object System.Drawing.Size(60,30)
$ok.Location			= New-Object System.Drawing.Point(290,140)
$ok.DialogResult		= [System.Windows.Forms.DialogResult]::OK
$ok.Text				= "OK"
$ok.TabIndex 			= 3
$ok.Enabled				= $False
	$form.Controls.Add($ok)
	
	$textBox1.Add_TextChanged({Check-Fill})
	$textBox2.Add_TextChanged({Check-Fill})

	$form.CancelButton	= $cancel
	$form.AcceptButton	= $ok

$fr = $form.ShowDialog()
If ($fr -eq "Cancel") { Break }
If ($fr -eq "OK") { 
	$script:group		= $textBox1.Text
	$script:userList	= $textBox2.Text
	Verification 
	}
}

Function Check-Fill {
	If ($textBox1.Text -AND $textBox2.Text) {
		$ok.Enabled	= $True
	} Else {
		$ok.Enabled	= $False
	}
}

Function Verification {
$groupList					= gam print groups |Select -Skip 1
$adUsers					= Get-ADUser -Filter * |Select SamAccountName
$googleUsers				= Gam Print Users |Select -Skip 1
$users						= $userList -creplace "\s" |% { $_.Split(",")}
$checkGroups				= $groupList |Select-String $group -Quiet
If ($checkGroups -eq $True) {
	ForEach ($user in $users) {
		$checkAD			= $adUsers |Select-String $user -Quiet
		$checkGoogle		= $googleUsers |Select-String $user -Quiet
		If (!$checkAD -OR !$checkGoogle) {
			[System.Windows.Forms.MessageBox]::Show("$user was not found in the system. Not added to the group." , "User Check" , "OK")
		} Else {
			gam update group $group add user $user
		}
	$script:updatedMembers	= gam info group $group |findstr " member: " |%{ $_.Replace(" member: ","")} | %{ $_ -creplace '@domain.*'} |Out-String
	}
} Else { 
		[System.Windows.Forms.MessageBox]::Show("The group you have entered was not found in the system. Please try again." , "Group Check" , "OK")
		Start-Form
	}
Show-Results
}

Function Show-Results {

$results					= New-Object System.Windows.Forms.Form
$results.Size				= New-Object System.Drawing.Size(280,300)
$results.Icon				= New-Object System.Drawing.Icon("C:\Scripts\Images\domain\domainLogo1.ico")
$results.Text				= "  Employment Tool"
$results.SizeGripStyle		= 'Hide'
$results.StartPosition		= 'CenterScreen'
$results.FormBorderStyle	= 'Fixed3D'
$results.Font				= "Arial,10,style=Bold"
$results.TopMost			= $True
$results.KeyPreview			= $True
$results.MaximizeBox		= $False
$results.MinimizeBox		= $False

$label						= New-Object System.Windows.Forms.Label
$label.Size					= New-Object System.Drawing.Size(350,40)
$label.Location				= New-Object System.Drawing.Point(20,5)
$label.Text					= "Updated Member List for: `n     $group"
	$results.Controls.Add($label)

$ok							= New-Object System.Windows.Forms.button
$ok.Size					= New-Object System.Drawing.Size(60,30)
$ok.Location				= New-Object System.Drawing.Point(190,230)
$ok.DialogResult			= [System.Windows.Forms.DialogResult]::OK
$ok.Text					= "OK"
$ok.TabIndex 				= 1
$ok.Enabled					= $True
	$results.Controls.Add($ok)	
		
$textBox1					= New-Object System.Windows.Forms.RichTextBox
$textBox1.Size				= New-Object System.Drawing.Size(200,170)
$textBox1.Location			= New-Object System.Drawing.Point(50,50)
$textBox1.TabIndex			= 0
$textBox1.Text				= "$updatedMembers"
$textBox1.ReadOnly			= $True
	$results.Controls.Add($textBox1)

$results.AcceptButton		= $ok
	
$fr		= $results.ShowDialog()
If ($fr -eq "Cancel") { Break }
If ($fr -eq "OK") { C:\Scripts\EmploymentTool_v3.0\GoogleGroups\Menu_GoogleGroups.ps1 }
}

Start-Form