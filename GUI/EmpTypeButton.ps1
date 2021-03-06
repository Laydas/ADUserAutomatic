[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "Data Entry Form"
$objForm.Size = New-Object System.Drawing.Size(450,200) 
$objForm.StartPosition = "CenterScreen"
$objForm.SizeGripStyle = "Hide"
$objForm.ShowInTaskbar = $False
$objForm.MinimizeBox = $False
$objForm.MaximizeBox = $False
$objForm.Topmost = $true
$Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$objForm.Icon = $Icon

$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$x=$objTextBox.Text;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})


$EmpButton = New-Object System.Windows.Forms.Button
$EmpButton.Location = New-Object System.Drawing.Size(75,120)
$EmpButton.Size = New-Object System.Drawing.Size(100,23)
$EmpButton.Text = "Employee"
$EmpButton.Add_Click({$EmpType="Employee";$objForm.Close()})
$objForm.Controls.Add($EmpButton)


$ConButton = New-Object System.Windows.Forms.Button
$ConButton.Location = New-Object System.Drawing.Size(175,120)
$ConButton.Size = New-Object System.Drawing.Size(100,23)
$ConButton.Text = "Contractor"
$ConButton.Add_Click({$EmpType="Contractor";$objForm.Close()})
$objForm.Controls.Add($ConButton)


$TempButton = New-Object System.Windows.Forms.Button
$TempButton.Location = New-Object System.Drawing.Size(275,120)
$TempButton.Size = New-Object System.Drawing.Size(100,23)
$TempButton.Text = "Temp/Student"
$TempButton.Add_Click({$EmpType="Temp";$objForm.Close()})
$objForm.Controls.Add($TempButton)


$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(10,20) 
$objLabel.Size = New-Object System.Drawing.Size(280,20) 
$objLabel.Text = "Please choose the type of account"
$objForm.Controls.Add($objLabel) 

$objForm.Topmost = $True

$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()

return $EmpType