#
#Remove after testing is complete.
Import-Module EmploymentTool_v3.0

Function Start-Form {
$form				= New-Form
$form.Size			= Set-Size 420 420
$form.Text			= "  Employment Tool"
$form.SizeGripStyle		= 'Hide'
$form.StartPosition		= 'CenterScreen'
$form.FormBorderStyle	= 'Fixed3D'
$form.Font			= "Arial,10,style=Bold"
$form.TopMost			= $True
$form.KeyPreview		= $True
$form.MaximizeBox		= $False
$form.MinimizeBox		= $False
$form.Icon			= Set-Icon

$fr = $form.ShowDialog()
}

Start-Form
