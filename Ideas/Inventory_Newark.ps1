# Employment Tool v3.0
# 	Inventory

$currentInventoryFile = "C:\Scripts\Text Files\Inventory - Newark.xml"
[xml]$currentInventory = Get-Content $currentInventoryFile
$MacbookAirs_Current = $currentInventory.inventory.macbookair
$MacbookPro13s_Current = $currentInventory.inventory.macbookpro13
$MacbookPro15s_Current = $currentInventory.inventory.macbookpro15

Function ShowInventory {
$form_Inventory = New-Form
$form_Inventory.Size = Set-Size 420 480
$form_Inventory.Text = "  Employment Tool - Inventory (Newark)"
$form_Inventory.SizeGripStyle = 'Hide'
$form_Inventory.StartPosition = 'CenterScreen'
$form_Inventory.FormBorderStyle = 'Fixed3D'
$form_Inventory.Font = "Arial,10,style=Bold"
$form_Inventory.TopMost = $True
$form_Inventory.KeyPreview = $True
$form_Inventory.MaximizeBox = $False
$form_Inventory.MinimizeBox = $False
$form_Inventory.Icon = Set-Icon

$label_MacbookAir = New-Label
$label_MacbookAir.Size = Set-Size 100 20
$label_MacbookAir.Location = Set-Point 10 10
$label_MacbookAir.Text = "Macbook Air"
	$form_Inventory.Controls.Add($label_MacbookAir)

$quantity_MacbookAir = New-TextBox
$quantity_MacbookAir.Enabled = $False
$quantity_MacbookAir.Size = Set-Size 30 20
$quantity_MacbookAir.Location = Set-Point 150 10
$quantity_MacbookAir.Text = $MacbookAirs_Current
	$form_Inventory.Controls.Add($quantity_MacbookAir)

	$button_MacbookAir_Decrease = New-Button
	$button_MacbookAir_Decrease.Text = "-"
	$button_MacbookAir_Decrease.Size = Set-Size 20 20
	$button_MacbookAir_Decrease.Location = Set-Point 120 10
	$button_MacbookAir_Decrease.Add_Click({
		[int]$n = $quantity_MacbookAir.Text
		$n -= 1
		$quantity_MacbookAir.Text = $n
		RV n
		})
		$form_Inventory.Controls.Add($button_MacbookAir_Decrease)

	$button_MacbookAir_Increase = New-Button
	$button_MacbookAir_Increase.Text = "+"
	$button_MacbookAir_Increase.Size = Set-Size 20 20
	$button_MacbookAir_Increase.Location = Set-Point 190 10
	$button_MacbookAir_Increase.Add_Click({
		[int]$n = $quantity_MacbookAir.Text
		$n += 1
		$quantity_MacbookAir.Text = $n
		RV n
		})
		$form_Inventory.Controls.Add($button_MacbookAir_Increase)

$label_MacbookPro13 = New-Label

$label_MacbookPro15 = New-Label

$button_Finish = New-Button
$button_Finish.Size = Set-Size 60 30
$button_Finish.Location = Set-Point 300 400
$button_Finish.DialogResult = [System.Windows.Forms.DialogResult]::OK
$button_Finish.Text = "Finish"
	$form_Inventory.Controls.Add($button_Finish)
	$form_Inventory.AcceptButton = $button_Finish

$fr = $form_Inventory.ShowDialog()
If ($fr -eq "OK") { $form_Inventory.Dispose() }
If ($fr -eq "Cancel") { $form_Inventory.Dispose() }
}

Function IncreaseNum ($param) {}

Function DecreaseNum ($param) {}

ShowInventory