############################################################
# TerminationAutomatic.ps1
# Written by Iain Letourneau
# Date last edited February 24th 2015
# Contact: Iain.Letourneau@gmail.com
#
# This script is custom written to run through all the
# pre-chosen Newalta termination tasks. This is done in 4
# steps all of which can run independantly of eachother.
#

# STEP 1 - IMPORT MODULES
# STEP 2 - GATHER INFORMATION
# STEP 3 - SET FUNCTIONS
# STEP 4 - TERMINATION STAGE 1
# STEP 5 - TERMINATION STAGE 2
# STEP 6 - TERMINATION STAGE 3
# STEP 7 - TERMINATION STAGE 4

############################################################


#===========================================================================
#======================== STEP 1 - IMPORT MODULES ==========================
#===========================================================================

#Import active directory to read current information
import-module ActiveDirectory
#Create Exchange session for quering user mailbox info
$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "PATH TO EXCHANGE SERVER"
Import-PSSession $ExchSession

# Change the directory so that the gui buttons will work
$location = Get-Item .
if($location.fullname -ne "PATH TO SCRIPT LOCATION")
{
    cd "PATH TO SCRIPT LOCATION"
}

# Load the GUI button
$GUIButton = (get-item .).fullname + "\AD User Automatic\GUI\CreateAnotherButton.ps1"
$CheckButton = (get-Item .).fullname + "\AD User Automatic\GUI\CheckBoxButton.ps1"
$LineEntry = (get-item .).fullname + "\AD User Automatic\GUI\LineEntry.ps1"
#===========================================================================

#Open excel object
$excel = new-object -comobject excel.application
#Open the excel layout sitting on my regular account desktop
$workbook = $excel.Workbooks.Open("PATH TO SCRIPTS\TerminateUser\_TerminatedUsers.xlsx")
#Select the first sheet
$ws = $workbook.WorkSheets.item(1)

#===========================================================================
#====================== STEP 3 - SET FUNCTIONS =============================
#===========================================================================

# Function for writing to excel
function WriteExcelStage() 
{
    #Open excel object and set as global variable
    $global:excelFun = new-object -comobject excel.application
    # Check to see if termination excel exists
    if((Test-Path $saveLoc) -eq $true)
    {
        $global:workbookFun = $excelFun.Workbooks.Open($saveLoc)
    }
    else
    {
        #Open the excel layout sitting on my regular account desktop
        $global:workbookFun = $excelFun.Workbooks.Add()
    }
    # Set interactive variables so I don't have to confirm saving the file
    #Select the first sheet
    $global:wsFun = $workbookFun.WorkSheets.item(1)
    #=============================================================================================================
}
# Initialize the function to close the excel file once written to
function ExitExcelStage() {
    # Save Excel
    if((Test-Path $saveLoc) -eq $true)
    {
        $workbookFun.Save()
    }
    else
    {
        $workbookFun.SaveAs($saveloc)
    }
    # Exit excel
    $workbookFun.Close()
    $ExcelFun.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsFun)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookFun)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelFun)
}
#===========================================================================

# Start loop to hit multiple terminations at once.
$AnotherUser = "Y"
do{
#===========================================================================
#===================== STEP 2 - GATHER INFORMATION =========================
#===========================================================================

#===================================================================
# GATHER INFORMATION NEEDED ON ALL STEPS

# Get the row of _TerminatedUsers.xlsx to read from
$row = . $LineEntry

if($exit -eq "Exit")
{
    $workbook.Close()
    $Excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
    Exit
}
# Pull data out of excel
$person = $ws.Cells.Item($row, 1).value2
$ticket = $ws.Cells.Item($row, 3).value2
$supervisor = $ws.Cells.Item($row, 2).value2
# Query AD for user information
$user = Get-ADUser $person -properties Description, EmailAddress, HomeDirectory, MemberOf, DistinguishedName
$supervisor = Get-ADUser $supervisor -properties HomeDirectory, EmailAddress
# Get the users name as first last
$firstLast = $user.GivenName + " " + $user.Surname

# Get all employees reporting to this person
# Initialize the reportString as empty
$reportString = ""
$reports = Get-ADUser $user -properties * | select -ExpandProperty directReports
if($reports -ne $null)
{
    foreach($report in $reports)
    {
        $reportSplit = $report.Split(",")
        $last = $reportSplit[0].substring(3,$reportSplit[0].length-4)
        $first = $reportSplit[1].trim()
        $reportString += $first + " " + $last + ", "
    }
    $reportString = $reportString.substring(0,$reportString.length-2)
}

# The relavent save location for this user
$saveloc = "PATH TO SCRIPTS\Excels\Termination " + $ticket + " - " + $firstLast + ".xlsx"

# Determine if the user is in the East or the West/South

if($user.HomeDirectory -like "*home3*")
{
    $eastWest = "East"
}
else
{
    $eastWest = "West"
}
$ButtonLabel = "$firstLast resides in $eastWest would you like to continue?`r`n"
$continue = . $GUIButton

if($continue -eq "N")
{
    continue
}
#--------------------------------------------------------------
# Get the date of removal, 30 days from now
$removalTime = (get-Date).AddDays(30)
$removalTime = get-Date $removalTime -f D
#--------------------------------------------------------------
#===================================================================




#==============================================================================================
#==============================================================================================
# = = = = = = = =                          STAGE 1                            = = = = = = = = = 
#==============================================================================================
#==============================================================================================
$WindowName = "STAGE 1"
$ButtonLabel = "Stage 1 will read the current status of the account in question. Once the statistics have"
$ButtonLabel += " been received the script will display the current status of the account and suggest"
$ButtonLabel += " which changes be made. If the user replies with yes then the changes are applied to"
$ButtonLabel += " the account and changes made are written to an excel file.`r`n`r`nDo you wish to run Stage 1 on this user?"
$stage1 = . $GUIButton
$WindowName = $null

if($stage1 -eq "Y"){

#==============================================================
# GET INFORMATION NEEDED FOR STAGE 1

#--------------------------------------------------------------
# Start compiling message for output to administrator
$ButtonLabel = "=== CURRENT AND PROPOSED CHANGES FOR " + $user.Name + " ===`r`n`r`n"
#--------------------------------------------------------------

#--------------------------------------------------------------
# Get all the group memberships
#--------------------------------------------------------------
# Initialize array for holding groups
$groups = @()
# Step through each group the user is a member of
foreach($group in $user.memberof)
{
    # Split the Distinguished name into something more readable
    $groupSplit = $group.split(","); 
    $groupName = $groupSplit[0].substring(3,$groupSplit[0].length-3)
    # If the group starts with \ such as the # groups remove the \
    if($groupName -match "^\\")
    {
        $groupName = $groupName.substring(1,$groupName.length-1)
    }
    # Add group to groups array
    $groups += $groupName
}
# Sort alphabetically
$groups = $groups | Sort
#--------------------------------------------------------------


#--------------------------------------------------------------
# check if enabled or not
if($user.enabled -eq $true)
{
    $userEnabled = $true
    $ButtonLabel += "Account Enabled: True ==> False`r`n"
}
else
{
    $userEnabled = $false
    $ButtonLabel += "Account Enabled: False`r`n"
}
#--------------------------------------------------------------

#--------------------------------------------------------------
# Check to see if the user has an email address
if($user.EmailAddress -ne $null)
{
    # Check if user is hidden from GAL
    $mailbox = Get-Mailbox $user.emailaddress
    if($mailbox.HiddenFromAddressListsEnabled -eq $true)
    {
        $userHidden = $true
        $ButtonLabel += "Hidden from GAL: True`r`n"
    }
    else
    {
        $userHidden = $false
        $ButtonLabel += "Hidden from GAL: False ==> True`r`n"
    }
}
else
{
    $userHidden = $true
    $ButtonLabel += "Hidden from GAL: No Mailbox found`r`n"
}
#--------------------------------------------------------------

#--------------------------------------------------------------
# Check to see if user has mailbox
if($user.EmailAddress -ne $null)
{
    # Check to see if ActiveSync is currently enabled
    $sync = get-CASMailbox -Identity $user.samaccountname
    if($sync.ActiveSyncEnabled -eq $true)
    {
        $userASync = $true
        $ButtonLabel += "ActiveSyncEnabled: True ==> False`r`n" 
    }
    else
    {
        $userASync = $false
        $ButtonLabel += "ActiveSyncEnabled: False`r`n"
    }
}
else
{
    $userASync = $false
    $ButtonLabel += "ActiveSyncEnabled: No mailbox found`r`n"
}
#--------------------------------------------------------------

#--------------------------------------------------------------
# Check to see if description has been tagged
$desc = $user.description
if($desc.length -gt 13)
{
    $tag = $desc.substring(0,14)
}
else
{
    $tag = " "
}
if($tag -like "TERM ?????? | ")
{
    $userTagged = $true
    $ButtonLabel += "Description Tag: " + $user.Description + "`r`n"
}
else
{
    $userTagged = $false
    $ButtonLabel += "New Description Tag: TERM " + $ticket + " | " + $user.Description + "`r`n"
    $termDesc = "TERM " + $ticket + " | " + $user.Description
}
#--------------------------------------------------------------
# END INFORMATION GATHERING/VARIABLE ASSIGING FOR STAGE 1
#===================================================================


# APPLY CHANGES TO AD ACCOUNT ====================================================================
# Ask if the proposed changes are to be applied
$makeChange = . $GUIButton

if($makeChange -eq "Y")
{
    #--------------------------------------------------------------
    # Open excel file for writing what was done
    WriteExcelStage
    #--------------------------------------------------------------
    
    
    #--------------------------------------------------------------
    # Disable Account
    #--------------------------------------------------------------
    if($userEnabled -eq $true)
    {
        Disable-AdAccount $user
        Write-Host "Account Disabled"
        $wsFun.Cells.Item(1,6) = "Account Disabled"
    }
    else
    {
        $wsFun.Cells.Item(1,6) = "Account Previously Disabled"
    }
    #--------------------------------------------------------------
    
    
    #--------------------------------------------------------------
    # Hide from GAL
    #--------------------------------------------------------------
    if($userHidden -eq $false)
    {
        Set-Mailbox $user.EmailAddress -HiddenFromAddressListsEnabled $true
        Write-Host "Hidden from GAL"
        $wsFun.Cells.Item(2,6) = "Account Hidden From GAL"
    }
    else
    {
        $wsFun.Cells.Item(2,6) = "Previously Hidden in GAL"
    }
    #--------------------------------------------------------------
    
    
    #--------------------------------------------------------------
    # Disable ActiveSync
    #--------------------------------------------------------------
    if($userASync -eq $true)
    {
        Set-CASMailbox -Identity $user.samaccountname -ActiveSyncEnabled $false
        Write-Host "Active Sync Disabled"
        $wsFun.Cells.Item(3,6) = "Disabled ActiveSync"
    }
    else
    {
        $wsFun.Cells.Item(3,6) = "ActiveSync Previously Disabled"
    }
    #--------------------------------------------------------------
    
    
    #--------------------------------------------------------------
    # Tag Description
    #--------------------------------------------------------------
    if($userTagged -eq $false)
    {
        Set-ADUser $user -description $termDesc
        Write-Host "Description Tagged"
        $wsFun.Cells.Item(4,6) = "Account Description Tagged"
    }
    else
    {
        $wsFun.Cells.Item(4,6) = "Description Already Tagged"
    }
    #--------------------------------------------------------------
    
    
    #--------------------------------------------------------------
    # Copy out Groups
    #--------------------------------------------------------------
    $row = 2
    $wsFun.Cells.Item(1,8) = "Group Memberships"
    foreach($group in $groups)
    {
        $wsFun.Cells.Item($row, 8) = $group
        $row++
    }
    $wsFun.Cells.Item(5,6) = "Copied out Groups"
    Write-Host "Groups placed in excel"
    #--------------------------------------------------------------
    
    
    #--------------------------------------------------------------
    # Remove Groups
    #--------------------------------------------------------------
    foreach($group in $groups)
    {
        $getGroup = get-adgroup $group
        Remove-ADGroupMember -Identity $getGroup -Member $user -Confirm:$false
    }
    $wsFun.Cells.Item(6,6) = "Removed Groups"
    Write-Host "All Groups have been removed from user"
    #--------------------------------------------------------------
    
    
    #--------------------------------------------------------------
    # Remove Phone Number
    #--------------------------------------------------------------
    set-aduser $user -OfficePhone $null -Fax $null -HomePhone $null -MobilePhone $null
    $wsFun.Cells.Item(7,6) = "Removed Phone Number"
    Write-Host "Phone numbers removed from Account"
    #--------------------------------------------------------------
    
    
    #--------------------------------------------------------------
    # Write stage1 complete and close excel
    #--------------------------------------------------------------
    $wsFun.Cells.Item(9,6) = "Stage 1 Complete"
    ExitExcelStage
    #--------------------------------------------------------------
}

}
#===================================================================


#==============================================================================================
#==============================================================================================
# = = = = = = = =                          STAGE 2                            = = = = = = = = = 
#==============================================================================================
#==============================================================================================

$WindowName = "STAGE 2"
$ButtonLabel = "Stage 2 will determine the  size of the home folder and also the number of files and folders"
$ButtonLabel += " within the home drive. The user's mailbox statistics will then be gathered and written"
$ButtonLabel += " to excel. The script will then compile the text for the email to send. The user is then"
$ButtonLabel += " moved to the terminated OU and PST folder created in the home drive of the user.`r`n`r`nDo you wish to run Stage 2 on this user?"
$stage2 = . $GUIButton
$WindowName = $null

if($stage2 -eq "Y"){

Write-Host "Stepping into stage2 of user termination script"

#==============================================================
# GET INFORMATION NEEDED FOR STAGE 2

#--------------------------------------------------------------
# Get all information needed from Exchange
#--------------------------------------------------------------
if($user.EmailAddress -ne $null)
{
Write-Host "Gathering information from Exchange"
# Query Exchange for mailbox info
$stats = Get-MailboxStatistics $user.emailaddress
# Turn the mail size into a readable format
$sizeSplit = $stats.TotalItemSize.split(" ")
$mailSize = $sizeSplit[0] + " " + $sizeSplit[1]
}
else
{
    Write-Host "User has no mailbox"
}
#--------------------------------------------------------------


#--------------------------------------------------------------
# Get all information needed from Active Directory
#--------------------------------------------------------------
Write-Host "Gathering information from Active Directory"
# Get the homedirectory path
$path = $user.HomeDirectory
#--------------------------------------------------------------


#--------------------------------------------------------------
# Get all information from system/servers
#--------------------------------------------------------------
Write-Host "Gathering general information"

#--------------------------------------------------------------
# Get the Home folder size, files, and folders
$files = [System.IO.Directory]::GetFiles($path, '*', 'AllDirectories').Count
$size = ((Get-ChildItem $path -Recurse) | Measure-Object -sum length).sum
$folders = [System.IO.Directory]::GetDirectories($path, '*', 'AllDirectories').Count

#--------------------------------------------------------------
# Turn the properties.size into a readable number
$homesize = "{0:N0}" -f $size + " Bytes"
if($size/1KB -ge 1)
{
    $homesize = "{0:N2}" -f ($size/1KB) + " KB"
}
if($size/1MB -ge 1)
{
    $homesize = "{0:N2}" -f ($size/1MB) + " MB"
}
if($size/1GB -ge 1)
{
    $homesize = "{0:N2}" -f ($size/1GB) + " GB"
}
#--------------------------------------------------------------

#--------------------------------------------------------------
# Create the path for creating the PST folder at
$pstFolder = $path + "\PST"
#--------------------------------------------------------------
#==============================================================



#==============================================================
# STAGE 2 ACTIONS

#--------------------------------------------------------------
# Compile text for email
#--------------------------------------------------------------
# Call function to open excel
Write-Host "Opening TERMINATEDUSERS template"
WriteExcelStage
#--------------------------------------------------------------

#--------------------------------------------------------------
# Write message and steps finished to the excel file
Write-Host "Writing to excel file"
$wsFun.Cells.Item(1,1) = "Hi " + $supervisor.GivenName + ","
$wsFun.Cells.Item(1,3) = "Ticket " + $ticket + " - Termination - " + $firstLast
$wsFun.Cells.ITem(1,4) = $supervisor.EmailAddress
if($user.EmailAddress -ne $null)
{
    $wsFun.Cells.Item(3,1) = $firstLast + "'s mailbox has " + $stats.ItemCount + " items, is a total of " + $mailSize + " in size and is located on Mail Server " + $stats.Database + "."
    $wsFun.Cells.Item(6,1) = "Do you wish for yourself or someone else to have access to this mailbox for a month before its removed from our system?"
    $wsFun.Cells.Item(7,1) = "Removal would occur on " + $removalTime + "."
}
else
{
    $wsFun.Cells.Item(3,1) = $firstLast + " has no mailbox."
}
$wsFun.Cells.Item(4,1) = "Their H: drive, located at " + $path + ", has " + $files + " Files, " + $folders + " Folders for a total size of " + $homeSize + "."

$wsFun.Cells.Item(9,1) = "Do you wish for yourself or someone else to have a copy of the H: drive to keep for as long as you wish?"
$wsFun.Cells.Item(10,1) = "Or may these be removed at IT's convenience?"
if($reports -ne $null)
{
    $wsFun.Cells.Item(12,1) = $firstLast + " is the manager for the following individuals: " + $reportString + "."
    $wsFun.Cells.Item(13,1) = "Could you please provide the new manager for those listed above."
    $wsFun.Cells.ITem(15,1) = "Thank you," 
}
else
{
    $wsFun.Cells.ITem(12,1) = "Thank you,"
}
$wsFun.Cells.Item(11,6) = "Moved Account to Disabled Container"
$wsFun.Cells.Item(12,6) = "Created PST folder in H: drive"
$wsFun.Cells.Item(13,6) = "Emailed Manager"
$wsFun.Cells.Item(15,6) = "Stage 2 Complete"
#--------------------------------------------------------------

#--------------------------------------------------------------
# Call function to close excel
Write-Host "Closing excel file"
ExitExcelStage
#--------------------------------------------------------------

#--------------------------------------------------------------
# Move Account to Disabled Container
Write-Host "Moving user to the Terminated Users OU"
Move-ADObject $user -TargetPath 'TERMINATED USERS AD OU DIRECTORY' #THIS LINE NEEDS TO BE UPDATED TO WORK
#--------------------------------------------------------------

#--------------------------------------------------------------
# Create PST Folder in H Drive
Write-Host "Creating PST folder in H Drive"
New-Item -ItemType directory -Path $pstFolder
#--------------------------------------------------------------

#--------------------------------------------------------------
Write-Host "End of Stage2 of Terminated Users Script"
}
#==============================================================




#==============================================================================================
#==============================================================================================
# = = = = = = = =                          STAGE 3                            = = = = = = = = = 
#==============================================================================================
#==============================================================================================

# Compile message to explain stage 4 to administrator running program
$WindowName = "STAGE 3"
$ButtonLabel = "Stage 3 will ask whether the supervisor requires the user's mailbox and H drive. If so"
$ButtonLabel += " full mailbox access will be granted to the supervisor and a copy of the H drive will be"
$ButtonLabel += " placed on the supervisor's H drive. The user's H drive will then be moved to the deleted"
$ButtonLabel += " home folders location..`r`n`r`nDo you wish to run Stage 3 on this user?"
$stage3 = . $GUIButton
$WindowName = $null

if($stage3 -eq "Y"){

# STAGE 3 ACTIONS
#--------------------------------------------------------------
$fileDate = (Get-Date (Get-Date (Get-Item $saveLoc).CreationTime).AddDays(30) -format D)

WriteExcelStage

$wsFun.Cells.Item(16, 1) = "Hi " + $supervisor.GivenName + ","
#--------------------------------------------------------------
# Grant mailbox access
if($user.EmailAddress -ne $null)
{
    $ButtonLabel = "Please check off which resources the manager has requested access to below."
    $mbxAccess = "N"
    $move = "N"
    $accessRequest = . $CheckButton
    if($accessRequest[0] -eq "Y")
    {
        Write-Host "Granting mailbox permissions to manager"
        Add-MailboxPermission -Identity $user.EmailAddress -User $supervisor.EmailAddress -AccessRights FullAccess -InheritanceType All
        $wsFun.Cells.Item(17,6) = "Granted Mailbox Access"
        $wsFun.Cells.Item(18,1) = "I have granted you access to the mailbox, it should be available the next time you open Outlook."
        $wsFun.Cells.Item(19,1) = "Unless I hear otherwise from you, this mailbox is scheduled to be  removed from our system on " + $filedate + "."
        $mbxaccess = "Y"
    }
    else
    {
        $wsFun.Cells.Item(17,6) = "Manager did not want the mailbox"
    }
}
else
{
    $wsFun.Cells.Item(17,6) = "User has no mailbox"
}
#--------------------------------------------------------------


#--------------------------------------------------------------
# Copy to Manager H drive

$superPath = $supervisor.HomeDirectory + "\" + $user.SamAccountName
# Check if there is a desktop.ini file and then remove it
$iniCheck = Get-ChildITem $user.HomeDirectory -force
foreach($inifile in $iniCheck)
{
    if($inifile.Name -eq "desktop.ini")
    {
        rm -force -path $inifile.FullName
    }
}
    
if($AccessRequest[1] -eq "Y")
{
    Write-Host "Copying H Drive to $superPath"
    robocopy $user.HomeDirectory $superPath /E

    $wsFun.Cells.Item(18,6) = "Copied H drive to Manager"
    $wsFun.Cells.Item(21,1) = "I have copied their H:\ drive to your H:\ drive. It is in a folder called " + $user.SamAccountName + ". This is your copy, you may edit or delete as you like."
    $move = "Y"
}
else
{
    $wsFun.Cells.Item(18,6) = "Manager did not want copy of H Drive"
}
$wsFun.Cells.Item(23,1) = "Regards,"
#--------------------------------------------------------------


#--------------------------------------------------------------
# Move to deleted home folders
$movePath = "PATH TO ARCHIVE OF USER DATA" + $user.SamAccountName
Write-Host "Moving H drive to Deleted Home Folders"
robocopy $user.HomeDirectory $movePath /E /MOVE
$wsFun.Cells.Item(19,6) = "Moved H to Deleted Users Folder"
#--------------------------------------------------------------

#--------------------------------------------------------------
if($move -eq "Y" -or $mbxAccess -eq "Y")
{
    $wsFun.Cells.Item(20,6) = "Notify Access Granted"
    $wsFun.Cells.Item(22,6) = "Stage 3 Complete"
    $wsFun.Cells.Item(23,6) = "Stage 4 Due " + $fileDate
}
else
{
    $wsFun.Cells.Item(20,6) = "No Access Given, Notification not neccessary"
    $wsFun.Cells.Item(22,6) = "Stage 3 Complete"
    $wsFun.Cells.Item(23,6) = "Stage 4 Due Immediately"
}
#--------------------------------------------------------------

ExitExcelStage

}
#==============================================================================================
#                                        STAGE 3 END
#==============================================================================================




#==============================================================================================
#==============================================================================================
# = = = = = = = =                          STAGE 4                            = = = = = = = = = 
#==============================================================================================
#==============================================================================================

# Compile message to explain stage 4 to administrator running program
$WindowName = "STAGE 4"
$ButtonLabel = "Stage 4 will create a mailbox export string and run it. The string will be saved to the excel for"
$ButtonLabel += " this user. Account Deletion is done manually.`r`n`r`nDo you wish to run Stage 4 on this user?"

  
if($user.EmailAddress -ne $null)
{
    # Get response from individual running program
    $stage4 = . $GUIButton
    $WindowName = $null
}
else
{
    $stage4 = "N"
    Write-Host "User has no mailbox"
    $wsFun.Cells.Item(25, 6) = "User has no Mailbox"
}
if($stage4 -eq "Y"){

# STAGE 4 GATHER INFORMATION
#--------------------------------------------------------------
# Get the export path of the pst file
$exportPath = "PATH TO TERMINATED USER DATA LOCATION" + $user.samaccountname + "\PST\" + $user.SamAccountName + ".pst"
$checkPstPath = "PATH TO TERMINATED USER DATA LOCATION" + $user.samaccountname + "\PST\"
Write-Host "Path for exporting the mailbox is $exportPath"
# Compile command output for reference
$command = "New-MailboxExportRequest -Mailbox " + $user.EmailAddress + " -FilePath " + $exportPath
#--------------------------------------------------------------

# STAGE 4 ACTIONS

#--------------------------------------------------------------
# Export the mailbox to the Deleted Users Folder
#--------------------------------------------------------------
if((Test-Path $checkPstPath) -eq $false)
{
    New-Item -Type directory -Path $checkPstPath
}

New-MailboxExportRequest –Mailbox $user.emailAddress –FilePath $exportPath
Write-Host "Mailbox has been exported"
#--------------------------------------------------------------


#--------------------------------------------------------------
# Write to excel the command used
#--------------------------------------------------------------
WriteExcelStage
$wsFun.Cells.Item(25, 1) = $command
$wsFun.Cells.Item(25, 6) = "Exported User's Mailbox"
ExitExcelStage
#--------------------------------------------------------------

}
#==============================================================================================
#                                      STAGE 4 END
#==============================================================================================


$WindowName = "Additional User"
$ButtonLabel = "Do you want to terminate another user?"
$AnotherUser = . $GUIButton
$WindowName = $null
}

while($AnotherUser -eq "Y")

$workbook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)

# Remove any left over variables that need to be removed
Write-Host "Removing Imported Exchange Session"
Remove-PSSession $ExchSession