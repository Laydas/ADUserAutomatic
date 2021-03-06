############################################################
# BiWeekly Increment HR Report Scan.ps1
# Written by Iain Letourneau
# Date last edited March 30 2016
# Contact: Iain.Letourneau@gmail.com
#
# The script will then check to see if 2 txt files exists HireText.txt and TermText.txt
# these two files are very important for saving large amount of time.
#
# The script will open the _TerminatedUsers.xlsx and the _NewHireUsers.xlsx files
# and then compare the txt files against the xlsx files, the txt files will only
# need to append any new terminations or new hire since last being run.
# It is these txt files that is used to verify if an account has already been
# terminated or hired yet.
#
# This script will check the location for the HR reports and
# find the newest report to work off of automatically.
# Naming convention must be as follows "IT Report - [Month] [Day] [year].xlsx"
#
# Once the newest file is found then the script will start stepping through
# each row
#
# The excel file needs to have specific columns in the following column
# Employee No. - A
# Action Type - C
# Manager's Name - F
# System User ID (UserID) - G
# Location - I
# (The script could be altered to check the column headers and work off of that in the future)
#
# This script will go through and check any manager changes or location changes
# Any changes made will be output to the following file convention
# "[Year]-[Month Number]-[Day] ChangesMade.txt"
############################################################


import-module activedirectory

#==============================================================================
# This entire section is to automatically select the newset HR report
#==============================================================================
$loc = Get-Item "PATH TO AD REPORTS\AD Reports"
$files = Get-Childitem $loc
$hash = @{}
foreach($file in $files)
{
    if($file.name -like "IT Report -*")
    {
        $filesplit = $file.name -split " - "
        $valArray = ($file.FullName, (Get-Date $filesplit[1].substring(0,$filesplit[1].length -5)))
        $hash.Add($file.Name, $valArray)
    }
}

# After this loop the newest file should be in selected
foreach($key in $hash.GetEnumerator())
{
    if($newest -eq $null)
    {
        $newest = $key.value[1]
        $newestFile = $key.value[0]
        continue
    }
    if($key.Value[1] -gt $newest)
    {
        $newest = $key.Value[1]
        $newestFile = $key.value[0]
    }
}
# Open up the excel for the newest file
$excel = new-object -comobject excel.application
#Open the excel layout sitting on my regular account desktop
$workbook = $excel.Workbooks.Open($newestFile)
#Select the first sheet
$ws = $workbook.WorkSheets.item(1)
$excel.visible = $true
#====================================================================================


#====================================================================================
#========================== SECTION FOR NEW/REHIRE USERS ============================
#====================================================================================

# Open up the new/rehire users Excel-----------------
$excelHire = new-object -comobject excel.application
$workbookHire = $excelHire.Workbooks.Open("PATH TO NEWHIREUSERS\AD User Automatic\_NewHireUsers.xls")
$wsHire = $workbookHire.WorkSheets.item(1)
$excelHire.visible = $true
#----------------------------------------------------

# Set the new/rehire text file location--------------
$HireTextLoc = "PATH TO HIRETEXT\HR Report Auto Check\HireText.txt"
#----------------------------------------------------

# Initialize the hiresarray--------------------------
$hiresArray = @()
#----------------------------------------------------

# Check to see if there is an existing new/rehire text file
if(Get-Item $hireTextLoc)
{
    Write-Host "Found an existing new/rehire text file, updating this should only take around 3 min"
    # Get the termination text file
    $HireFile = Get-Content $hireTextLoc
    # Assign text into array
    foreach($line in $hireFile) 
    {
        $hiresArray += $line
    }
    # Step through the termiantion excel starting at the last location.
    for($rowHire = $hiresArray.length + 7; $rowHire -lt $wsHire.UsedRange.Rows.Count; $rowHire++)
    {
        if($wsHire.Cells.Item($rowHire, 9).text -eq "")
        {
            break
        }
        $hiresArray += $wsHire.Cells.Item($rowHire, 9).text
    }
    # Make sure to save and update the text file
    $hiresArray > $hireTextLoc
}
else
{
    Write-Host "No new/rehire text file found, creating the first one, please wait as this can take 20min or longer depending on size"
    for($rowHire = 7; $rowHire -lt $wsHire.UsedRange.Rows.Count; $rowHire++)
    {
        if($wsHire.cells.item($rowHire, 9).text -eq "")
        {
            break
        }
        $hiresArray += $wsHire.Cells.Item($rowHire, 9).text
    }
    # SAve to termination text location for next time
    $hiresArray > $hireTextLoc
}
#----------------------------------------------------
Write-Host "Completed compiling new hires list"

# Remove all variables and excel related to the termination excel
$ExcelHire.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsHire)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookHire)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelHire)
Remove-Variable wsHire
Remove-Variable workbookHire
Remove-Variable ExcelHire
# END NEW/REHIRE EXCEL READING SECTION
#====================================================================================



#====================================================================================
#========================== SECTION FOR TERMINATED USERS ============================
#====================================================================================

# Open up the terminated users Excel-----------------
$excelTerm = new-object -comobject excel.application
$workbookTerm = $excelTerm.Workbooks.Open("PATH TO SCRIPTS\TerminateUser\_TerminatedUsers.xlsx")
$wsTerm = $workbookTerm.Worksheets.item(1)
#----------------------------------------------------

# Set the termination text file location-------------
$termTextLoc = "PATH TO SCRIPTS\HR Report Auto Check\TermText.txt"
#----------------------------------------------------

# Initialiaze the termsarray-------------------------
$termsArray = @()
#----------------------------------------------------

# Check to see if there is an existing termination text file
if(Get-Item $termTextLoc)
{
    Write-Host "Found an existing termination text file, updating this should only take around 1 min"
    # Get the termination text file
    $TermFile = Get-Content $termTextLoc
    # Assign text into array
    foreach($line in $termFile) 
    {
        $termsArray += $line
    }
    # Step through the termiantion excel starting at the last location.
    for($rowTerm = $termsArray.length + 3; $rowTerm -lt $wsTerm.UsedRange.Rows.Count; $rowTerm++)
    {
        if($wsTerm.Cells.Item($rowTerm, 1).text -eq "")
        {
            break
        }
        $termsArray += $wsTerm.Cells.Item($rowTerm, 1).text
    }
    # Make sure to save and update the text file
    $termsArray > $termTextLoc
}
# If the termination text file doesn't exist then create it.
else
{
    Write-Host "No termination text file found, creating the first one, please wait as this can take 10min or longer depending on size"
    for($rowTerm = 3; $rowTerm -lt $wsTerm.UsedRange.Rows.Count; $rowTerm++)
    {
        if($wsterm.cells.item($rowTerm, 1).text -eq "")
        {
            break
        }
        $termsArray += $wsTerm.Cells.Item($rowTerm, 1).text
    }
    # SAve to termination text location for next time
    $termsArray > $termTextLoc
}
#----------------------------------------------------
Write-Host "Completed compiling terminated users list"

# Remove all variables and excel related to the termination excel
$ExcelTerm.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsTerm)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookTerm)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelTerm)
Remove-Variable wsTerm
Remove-Variable workbookTerm
Remove-Variable ExcelTerm
    # Assign the location into a variable
    $loc = $ws.Cells.Item($row, 9).Text
    # Assign the manager's name into a variable
# END TERMINATION EXCEL READING SECTION
#====================================================================================

$txtout = ""
#====================================================================================
#=========== APPLY ACTIONS TO HR REPORT BASED ON COLLECTED INFORMATION ==============
#====================================================================================
for($row = 2; $row -lt $ws.UsedRange.Rows.Count; $row++)
{
    Write-Host "At row $row"
    $txt += "-----------------------Row: $row------------------------`r`n"
    $found = $false
    Write-Host "Getting information from HR Report"
    # Assign the samaccountname into a variable
    $sam = $ws.Cells.Item($row, 7).Text
    # Assign the employee number into a variable
    $empNum = $ws.Cells.Item($row, 1).Text
    $empMan = $ws.Cells.Item($row, 6).Text
    # Split the manager name
    $ManSplit = $empMan.split(" ")
    # Get the first and last names into variables
    $first = $ManSplit[0].trim()
    $last = $ManSplit[1].trim()
    # Compile searchable string
    $ManSearch = $last + ", " + $first
   
    # Check to see if the SAMACcountName is filled out
    if($sam -ne "")
    {
        if(dsquery user -samid $sam)
        {
            $user = Get-ADUser $sam -properties EmployeeNumber, Office, Manager
            Write-Host "Found user in AD:" $user.name
            $txt += "User found in AD: " + $user.Name + "`r`n"
        }
    }
    else
    {
        Write-Host "Did not find in AD, skipping to next row"
        $txt += "User not found in AD, skipping to next row`r`n"
        $ws.Range("A$($row)`:K$($row)").interior.colorindex = 6
        continue
    }
    
    Write-Host "...Searching for manager in AD"
    # Check to see if the manager exists/can be found
    if(dsquery user -name $ManSearch)
    {
        # Get the manager by first and last name
        $Manager = Get-ADUser -filter {givenName -eq $first -and surname -eq $last}
        # If 3 or more managers are found then mark as null and continue
        if($manager.length -gt 2)
        {
            $manager = $null
            continue
        }
        # If 2 managers are found check to see if one is admin account and remove it
        if($manager.length -eq 2)
        {
            for($i = 0; $i -lt $Manager.length; $i++)
            {
                if($Manager[$i].SAMAccountName.length -eq 2 -or $manager[$i].SamAccountName -like "[A-Z][A-Z][0-9]")
                {
                    $Manager = $Manager[($i+1)%2]
                    break
                }
            }
        }
        # If no admin account was caught then mark as null
        if($manager.length -eq 2)
        {
            $manager = $null
        }
    }
    else
    {
        $Manager = $null
    }
    Write-Host "... Manager search complete"
    
    # Verify against current data and update
    if($user.Office -ne $loc)
    {
        Write-Host "Updating office location for " $user.Name "`r`nOld Office: " $user.Office "`r`nNew Office: " $loc
        $txt += "Updating office location - Old Office: " + $user.Office + " - New Office: " + $loc + "`r`n"
        Set-ADUser $user -Office $loc
    }
    if($user.EmployeeNumber -ne $empNum)
    {
        Write-Host "Updating employee number for " $user.Name "`r`nOld Number: " $user.EmployeeNumber "`r`nNew Number: " $empNum
        $txt += "Updating employee number - Old Number: " + $user.EmployeeNumber + " - New Number: " + $empNum + "`r`n"
        Set-ADUser $user -EmployeeNumber $empNum
    }
    if($manager -ne $null)
    {
        if($user.Manager -ne $manager.distinguishedName)
        {
            Write-Host "Updating manager for " $user.Name "`r`nOld manager: " $user.Manager "`r`nNew Manager: " $manager.DistinguishedName
            $txt += "Updating manager - Old Manager: " + $user.Manager + " - New Manager: " + $manager.DistinguishedName + "`r`n"
            Set-ADUser $user -Manager $Manager.DistinguishedName
        }
    }
    
    Write-Host "Checking if user exists in termination/newhire lists"
    switch($ws.Cells.ITem($row, 3).text)
    {
        "Termination" 
        {
            foreach($term in $termsArray)
            {
                if($term -eq $sam)
                {
                    Write-Host "Found in term list"
                    $txt += "Verified that this user has been terminated`r`n"
                    $found = $true
                    break
                }
            }
            # If the user was found in the terminated excel then highlight green, else highlight yellow
            if($found -eq $true) {$ws.Range("A$($row)`:K$($row)").interior.colorindex = 4}
            else {$ws.Range("A$($row)`:K$($row)").interior.colorindex = 6}
            
        }
        {($_ -eq "Hiring") -or ($_ -eq "rehire")} 
        {
            foreach($hire in $hiresArray)
            {
                if($hire -eq $sam)
                {
                    Write-Host "Found in hire list"
                    $txt += "Verified that this user was in the new hire list`r`n"
                    $found = $true
                }
            }
            # If the user was found in the new/rehire excel then highlight green, else highlight yellow
            if($found -eq $true) {$ws.Range("A$($row)`:K$($row)").interior.colorindex = 4}
            else {$ws.Range("A$($row)`:K$($row)").interior.colorindex = 6}
        }
    }
}
#====================================================================================

# Save the file
$workbook.Save()

$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
Remove-Variable ws
Remove-Variable workbook
Remove-Variable Excel

$date = (Get-Date -f yyyy/M/dd)
$date = $date.replace("/","-")
$fileloc = "PATH TO SERVER\AD Reports\ChangeLog\" + $date + " ChangesMade.txt"
$txt > $fileloc