############################################################
# Memberships.ps1
# Written by Iain Letourneau
# Date last edited Dec 8th 2014
# Contact: Iain.Letourneau@gmail.com
#
# This script meant to be used a function within CreateUser.ps1
# This script can be run on it's own as well for position changes
#
############################################################

#-----------------------------------------------------------
#Create a new excel object
$excelMirror = new-object -comobject excel.application
#Open the excel layout sitting on my regular account desktop
$workbookMirror = $excelMirror.Workbooks.Open("PATH TO SCRIPTS\scripts\ad user automatic\Functions\_LayoutCopy.xlsx")
#Select the first sheet
$wsMir = $workbookMirror.WorkSheets.item(1)
#-----------------------------------------------------------

#-----------------------------------------------------------
# Check to see if this is being run from CreateUser or not
if($setVar -eq "SET")
{
    # Gather all variables from CreateUser
    $fullname = $first + " " + $last
    $read1 = $fullname
    $read2 = $mirrorUser
    $readtick = $ticket
    $readreq = $Requestor
    
}
else
{
    #Import the Active Directory Module for querying data
    import-module activedirectory
    #Get the user information
    $read1 = Read-Host "Enter user getting permissions EX: Iain Letourneau"
    $read2 = Read-Host "Enter user to be mirrored EX: Iain Letourneau"
    $readtick = Read-Host "Enter Ticket # EX: 123456"
    $readreq = Read-Host "Enter Requesting User EX:Iain Letourneau"
}
#-----------------------------------------------------------

#-----------------------------------------------------------
# Get the SAMAccountName of the target user

# Grab the variable from CreateUser
if($setVar -eq "SET")
{
    $userSAM = $samaccountname
}
else
{
    #Split the first and last names of the source and target
    if($read1.split(" ").length -le 2)
    {
        $userfirst = $read1.split(" ")[0]
        $userlast = $read1.split(" ")[1]
    }
    else
    {
        # Could not find anything, have user enter it directly
        $userSAM = Read-Host "Please specify the user's SAMAccountName: "
    }    
}
#-----------------------------------------------------------

#-----------------------------------------------------------
#Do the same as above for the source user
if($read2.split(" ").length -le 2)
{
    $mirrorFirst = $read2.split(" ")[0]
    $mirrorLast = $read2.split(" ")[1]
}
else
{
    $mirrorSAM = Read-Host "Please specify the mirror user's SAMAccountName: "  
}
#-----------------------------------------------------------

#-----------------------------------------------------------
if($userSAM -eq $null)
{
    #Make sure the users exist before continuing-----------------------------
    $UserTest = Get-Aduser -Filter {givenname -eq $userFirst -and surname -eq $userLast}
    if($UserTest -eq $Null)
    {
        $userSAM = Read-Host "Could not find user, please specify SAMAccountName of the user"
    }
    if($UserTest.length -ge 2)
    {
        $userSAM = Read-Host "Found more than 1 user, please specify SAMAccountName of the user"
    }
    else
    {
        $userSAM = $userTest.SAMAccountName
    }
}
#-----------------------------------------------------------

#-----------------------------------------------------------
$UserTest = Get-ADuser -Filter {givenname -eq $mirrorFirst -and surname -eq $mirrorLast}
$mirrorSAM = $usertest.SAMAccountName

if($mirrorSAM -eq $null)
{
    
    if($UserTest -eq $Null)
    {
        $mirrorSAM = Read-Host "Could not find user, please enter in exact SamAccountName now"
    }
    if($userTest.length -ge 2)
    {
        $mirrorSAM = Read-Host "Found more than 1 user, please specify SAMAccountName of the mirror"
    }#-----------------------------------------------------------------------
}


#-----------------------------------------------------------

#-----------------------------------------------------------
#Get all the groups the users belong to
$target = Get-AdUser $userSAM -Properties memberof, description
$source = Get-AdUser $mirrorSAM -Properties memberof, description
#-----------------------------------------------------------

#-----------------------------------------------------------
#Get the description of the users (Job Title)
$targetdesc = $target.description
$sourcedesc = $source.description
#-----------------------------------------------------------

#-----------------------------------------------------------
#Set the excel file to start on line 2 after the headers
$line = 2
foreach($group in $source.memberof)
{
    #Reset the boolean
    $x = 0
    foreach($group2 in $target.memberof)
    {
        if($group -eq $group2)
        {
            $x = 1
        }
    }
    if($x -eq 0)
    {
        #only output the info to excel if 2nd user does not have the group
        $gname = get-adgroup $group
        $gdesc = get-adgroup $group -properties description   
        $ginfo = get-adgroup $group -properties info    
        $wsMir.Cells.Item($line,1) = $gname.Name
        $wsMir.Cells.Item($line,2) = $gdesc.description
        $wsMir.Cells.Item($line,3) = $ginfo.info
        $line++
    }
}

$wsMir.Cells.Item(2,7) = "Ticket " + $readtick + ": Folder Access Request for " + $read1 + " same as " + $read2 + " as per " + $readreq
$wsMir.Cells.Item(6,7) = $readreq + " has requested that " + $read1 + ", " + $targetdesc + " be granted the same permissions as " + $read2 + ", " + $sourcedesc + "."
$wsMir.Cells.Item(8,7) = "Please review the permissions requested next to your name as per below and reply with either APPROVE or DENY."
$wsMir.Cells.ITem(9,7) = "To keep volume of emails down please do not Reply All unless necessary."
#-----------------------------------------------------------

#-----------------------------------------------------------
# Save the excel file to the Excels Location
$saveloc = "PATH TO SCRIPTS\scripts\Excels\Ticket " + $readtick + " - " + $read1 + " same as " + $read2 + " request by " + $readreq + ".xlsx"
#-----------------------------------------------------------

#-----------------------------------------------------------
# Save the excel and then quit it
$workbookMirror.SaveAs($saveloc)

$workbookMirror.Close()
$ExcelMirror.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wsMir)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookMirror)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelMirror)
Remove-Variable wsMir
Remove-Variable workbookMirror
Remove-Variable ExcelMirror
#-----------------------------------------------------------