# Google Groups

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Function Start-Form {
$form					= New-Object System.Windows.Forms.Form
$form.Text				= "  Employment Tool"
$form.Size				= New-Object System.Drawing.Size(240,180)
$form.SizeGripStyle		= 'Hide'
$form.StartPosition		= 'CenterScreen'
$form.FormBorderStyle	= 'Fixed3D'
$form.Font				= "Arial,10,style=Bold"
$form.TopMost			= $True
$form.KeyPreview		= $True
$form.MaximizeBox		= $False
$form.MinimizeBox		= $False
$form.Icon				= New-Object System.Drawing.Icon("C:\Scripts\Images\SevOne\SevOneLogo1.ico")

$label					= New-Object System.Windows.Forms.Label
$label.Text				= "Choose A Category:"
$label.AutoSize			= $True
$label.Location			= New-Object System.Drawing.Point(10,10)
	$form.Controls.Add($label)
	
$listBox				= New-Object System.Windows.Forms.ListBox
$listBox.TabIndex 		= 0
$listBox.AutoSize		= $True
$listBox.Location		= New-Object System.Drawing.Point(30,30)
$listBox.BorderStyle	= 'FixedSingle'
$listBoxOptions			= "Create Group" , "Delete Group" , "Add/Remove Members" , "Modify Settings"
ForEach ($choice in $listBoxOptions) {
	$listBox.Items.Add($choice)
}
	$form.Controls.Add($listBox)

$cancel					= New-Object System.Windows.Forms.Button
$cancel.TabIndex 		= 1
$cancel.Text			= "Cancel"
$cancel.Size			= New-Object System.Drawing.Size(60,30)
$cancel.Location		= New-Object System.Drawing.Point(95,120)
$cancel.DialogResult	= [System.Windows.Forms.DialogResult]::Cancel
	$form.Controls.Add($cancel)
	
$ok						= New-Object System.Windows.Forms.button
$ok.TabIndex 			= 2
$ok.Text				= "OK"
$ok.Size				= New-Object System.Drawing.Size(60,30)
$ok.Location			= New-Object System.Drawing.Point(160,120)
$ok.DialogResult		= [System.Windows.Forms.DialogResult]::OK
$ok.Enabled				= $False
	$form.Controls.Add($ok)
	$listBox.Add_SelectedIndexChanged({$ok.Enabled	= $True})

	$form.CancelButton	= $cancel
	$form.AcceptButton	= $ok
	$fr = $form.ShowDialog()
	If ($fr -eq "Cancel"){C:\Scripts\EmploymentTool_v3.0\Menu_Main.ps1 ; $form.Dispose() ; Break }
	$r = $listBox.SelectedIndex
	Switch ($r) {
	'0'{C:\Scripts\EmploymentTool_v3.0\GoogleGroups\CreateGroup_GoogleGroups.ps1}
	'1'{C:\Scripts\EmploymentTool_v3.0\GoogleGroups\DeleteGroup_GoogleGroups.ps1}
	'2'{C:\Scripts\EmploymentTool_v3.0\GoogleGroups\Modifications_GoogleGroups.ps1}
	'3'{C:\Scripts\EmploymentTool_v3.0\GoogleGroups\ModifySettings_GoogleGroups.ps1}
	}
$form.Dispose()
}

Start-Form