[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
if($WindowName -ne $null)
{
    $objForm.Text = $WindowName
}
else
{
    $objForm.Text = "Data Entry Form"
}
$objForm.Size = New-Object System.Drawing.Size(450,280) 
$objForm.StartPosition = "CenterScreen"
$objForm.SizeGripStyle = "Hide"
$objForm.ShowInTaskbar = $False
$objForm.MinimizeBox = $False
$objForm.MaximizeBox = $False
$objForm.Topmost = $true
$font = New-Object System.Drawing.Font("Times New Roman",13,[system.drawing.fontstyle]::regular)
$objForm.Font = $font

$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$objForm.Icon = $Icon

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$objCheckBox1 = New-Object System.Windows.Forms.Checkbox
$objCheckBox1.Location = New-Object System.Drawing.Size(175,100)
$objCheckBox1.Size = New-Object System.Drawing.Size(250,23)
$objCheckBox1.Text = "Mailbox"
#$objCheckBox.Add_Click({$choice="Y";$objForm.Close()})
$objForm.Controls.Add($objCheckBox1)

$objCheckbox2 = New-Object System.Windows.Forms.Checkbox
$objCheckbox2.Location = New-Object System.Drawing.Size(175,130)
$objCheckbox2.Size = New-Object System.Drawing.Size(250,23)
$objCheckbox2.Text = "HomeDrive"
#$NoButton.Add_Click({$choice="N";$objForm.Close()})
$objForm.Controls.Add($objCheckbox2)

$SubmitButton = New-Object System.Windows.Forms.Button
$SubmitButton.Location = New-Object System.Drawing.Size(175,220)
$SubmitButton.Size = New-Object System.Drawing.Size(75,23)
$SubmitButton.Text = "Submit"
$SubmitButton.Add_Click({$objForm.Close()})
$objForm.Controls.Add($SubmitButton)


$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(425,200) 
$objLabel.Text = $ButtonLabel
$objForm.Controls.Add($objLabel) 

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

if($objCheckBox1.Checked -eq $true)
{
    $check1 = "Y"
}
else
{
    $check1 = "N"
}

if($objCheckBox2.Checked -eq $true)
{
    $check2 = "Y"
}
else
{
    $check2 = "N"
}

$choice = ($check1, $check2)

return $choice