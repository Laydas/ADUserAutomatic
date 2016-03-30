############################################################
# CreateUser.ps1
# Written by Iain Letourneau
# Date last edited March 30 2016
# Contact: Iain.Letourneau@gmail.com
#
# This script is custom written to create a new AD user
# based on the new hire document written by Jason Sallay
# this script will call multiple other neccessary scripts
# to properly work.
############################################################


#Import Active Directory-----------------------------------------------------------------------------------------------------------
import-module activedirectory
#----------------------------------------------------------------------------------------------------------------------------------


#Connect to Exchange Server--------------------------------------------------------------------------------------------------------
Write-Host "Connecting to exchange server"
#Create a connection to the Exchange server in order to enable the new user's email account
$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "PATH TO EXCHANGE SERVER"
#Open up the connection
Import-PSSession $ExchSession
#----------------------------------------------------------------------------------------------------------------------------------

# Change the location to where it needs to be for this script to run
$location = Get-Item .
if($location.fullname -ne "PATH TO SCRIPTS\AD User Automatic")
{
    cd "PATH TO SCRIPTS\AD User Automatic"
}
# Compile the strings for calling other functions
$autoEmail = (get-item .).fullname + "\Email Automatic\NewHireEmail.ps1"
$anotherUserButton = (get-item .).fullname + "\GUI\CreateAnotherButton.ps1"
$tempButtonPath = (Get-Item .).fullname + "\GUI\TemplateButton.ps1"
$empButtonPath = (Get-Item .).fullname + "\GUI\EmpTypeButton.ps1"
$lineEntry = (get-item .).fullname + "\GUI\LineEntry.ps1"
#----------------------------------------------------------------------------------------------------------------------------------


#Set variable for calling Memberships.ps1 so it wont interactively ask for information---------------------------------------------
$setVar = "SET"

#Open excel object
$excel = new-object -comobject excel.application
#Open the excel layout sitting on my regular account desktop
$workbook = $excel.Workbooks.Open("PATH TO SCRIPTS\AD User Automatic\_NewHireUsers.xls")
#Select the first sheet
$ws = $workbook.WorkSheets.item(1)
#----------------------------------------------------------------------------------------------------------------------------------

$UserCreatedList = @()

$AnotherUser = "Y"
do{
#============================================================================
# GATHER DATA FOR USER CREATION
#============================================================================
#Get info from user for template and excel
$row = . $lineEntry
if($exit -eq "Exit")
{
    Exit
}
$check = $ws.Cells.Item($row, 9).value2
if(dsquery user -samid $check){Write-Host "Username already found, Program will exit."; Start-Sleep -s 10; Exit}
#Get the location of the user
$template = . $tempButtonPath
#Step through the 5 options for templates
$EmpType = . $empButtonPath

switch ($template)
{
    "Corp" 
    {
        $lyncLoc = "West"
        switch($EmpType)
        {
            "Employee" {$templateUser = "corptemplateemp"}
            "Temp" {$templateUser = "corptemplateemp"}
            "Contractor" {$templateUser = "corpcontemplate"}
        }
    }
    "West" 
    {
        $lyncLoc = "West"
        switch ($EmpType)
        {
            "Employee" {$templateUser = "westtemplateemp"}
            "Temp" {$templateUser = "westtemplateemp"}
            "Contractor" {$templateUser = "templatewestcon"}
        }
    }
    "East" 
    {
        $lyncLoc = "East"
        switch($EmpType)
        {
            "Employee" {$templateUser = "buremotemplate"}
            "Temp" {$templateUser = "buremotemplate"}
            "Contractor" {$templateUser = "burcontemplate"}
        }
    }
    "French" 
    {
        $lyncLoc = "East"
        $templateUser = "eastfrenchtemp"
    }
    "South" 
    {
        $lyncLoc = "West"
        $templateUser = "texastemplateemp"
    }
}

#----------------------------------------------------------------------------

#Get the required information for creating new user
$first = $ws.Cells.Item($row, 3).value2
$last = $ws.Cells.Item($row, 4).value2
$displayName = $ws.Cells.Item($row, 5).value2
$SAMAccountName = $ws.Cells.Item($row, 9).value2
$Password = $ws.Cells.Item(1,3).value2
$Description = $ws.Cells.Item($row, 6).value2
$Office = $ws.Cells.Item($row, 7).value2
$ProfilePath = $ws.Cells.Item(3,3).value2
$RemoteHome = $ws.Cells.Item(3,6).value2
$Email = $ws.Cells.Item($row, 11).value2
$ExpiryDate = $ws.Cells.Item($row, 14).Text
$mirrorUser = $ws.Cells.Item($row, 16).value2
$Requestor = $ws.Cells.Item($row, 8).value2
$ticket = $ws.Cells.Item($row, 2).value2

#Set the correct home drive based on the location
if($template -eq "Corp" -or $template -eq "West")
{$homedir = "PATH TO USER DIRECTORY\"+$samaccountname; $EmailLoc = "WestCan"}

if($template -eq "East")
{$homedir = "PATH TO USER DIRECTORY\" + $samaccountname; $EmailLoc = "Atlantic"}

if($template -eq "French")
{$homedir = "PATH TO USER DIRECTORY\" + $samaccountname; $EmailLoc = "Quebec"}

if($template -eq "South")
{$homedir = "PATH TO USER DIRECTORY\"+$samaccountname; $EmailLoc = "South"}

# If the user exists at a remote site then direct home drive to correct location
switch -regex ($Office)
{
    #East Users
    "Barrie" {$EmailLoc = "Ontario"}
    "Bedford" {$EmailLoc = "Atlantic"}
    "Brantford" {$EmailLoc = "Ontario"}
    "Burlington" {$EmailLoc = "Ontario"}
    "Dartmouth" {$EmailLoc = "Atlantic"; $homedir = "PATH TO USER DIRECTORY\Dartmouth\Users\" + $samaccountname}
    "Erie" {$EmailLoc = "Ontario"}
    "Foxtrap" {$EmailLoc = "Atlantic"}
    "Hamilton" {$EmailLoc = "Ontario"}
    "Rexdale" {$EmailLoc = "Ontario"}
    "Toronto" {$EmailLoc = "Ontario"}
    "Sarnia" {$EmailLoc = "Ontario"}
    "St John" {$EmailLoc = "Atlantic"}
    "Stoney" {$EmailLoc = "Ontario"}
    "Sussex" {$EmailLoc = "Atlantic"}
    #French Users
    "Brossard" {$homedir = "PATH TO USER DIRECTORY\Brossard\Users\" + $samaccountname}
    "Chateauguay" {$homedir = "PATH TO USER DIRECTORY\Chateauguay\Users\" + $samaccountname}
    "VSC" {$homedir = "PATH TO USER DIRECTORY\VSC\Users\"+$samaccountname}
    #South Users
    "Denver" {$homedir = "PATH TO USER DIRECTORY\Denver\Users\"+$samaccountname}
    "Texas" {$homedir = "PATH TO USER DIRECTORY\Texas\Users\"+$samaccountname}
    #West Can Users
    "Edmonton" {$homedir = "PATH TO USER DIRECTORY\Edmonton\Users\"+$samaccountname}
    "Grand" {$homedir = "PATH TO USER DIRECTORY\GrandPrairie\Users\" + $samaccountname}
    "Leduc" {$homedir = "PATH TO USER DIRECTORY\Leduc\Users\"+$samaccountname}
    "NorthVan" {$homedir = "PATH TO USER DIRECTORY\NorthVan\Users\"+$samaccountname}
    "RedDeer" {$homedir = "PATH TO USER DIRECTORY\RedDeer\Users\"+$samaccountname}
    
}

#Create the folder location
New-Item -ItemType directory -Path $homedir

#Place all new users in the Root of Corp to make it easier to located and place in correct OU
$createPath = "AD OU PATH FOR NEWLY CREATED USERS"
$usrPName = $samaccountname
$usrPName += "@example.com"

#--------------------------------------------------------------------------------
#Create the new user
if($ExpiryDate)
{
    $ExpiryDate = (get-Date $ExpiryDate).AddDays(1)
    New-ADUser -SamAccountName $SAMAccountName -GivenName $first -Surname $last -Name $displayName -DisplayName $displayName -UserPrincipalName $usrPName -AccountPassword (ConvertTo-SecureString -AsPlainText $Password -force) -Office $Office -Path $createPath -AccountExpirationDate $ExpiryDate -Description $Description -HomeDirectory $homedir -HomeDrive "H:" -ScriptPath "Login.cmd"
}
else
{
    New-ADUser -SamAccountName $SAMAccountName -GivenName $first -Surname $last -Name $displayName -DisplayName $displayName -UserPrincipalName $usrPName -AccountPassword (ConvertTo-SecureString -AsPlainText $Password -force) -Office $Office -Path $createPath -Description $Description -HomeDirectory $homedir -HomeDrive "H:" -ScriptPath "Login.cmd"
}
#--------------------------------------------------------------------------------


Start-Sleep -s 5



#============================================================================
# STEPS TO COMPLETE ONCE THE USER IS CREATED
#============================================================================
# 1. Give user modify access to their home folder
# 2. Force change password at next logon
# 3. Set the TsProfilePath
# 3.1 Other Account Settings
# 4. Enable Account
Write-Host "Stepping into post-creation section"
#----------------------------------------------------------------------------
# 1. Give user modify access to their home folder
#----------------------------------------------------------------------------
# Set the until loop variable
$aclSet = $false
# Silence errors for this section of the script
$ErrorActionPreference = 'silentlyContinue'
# Get the homedir into an object
$homedir = Get-Item $homedir
# Start loop to ensure that the permissions are applied to the home folder
do
{
# Get the ACL on the object
$homeacl = Get-Acl $homedir
# Create the rule for giving the newly created user access to his home folder
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule($samaccountname,"Modify", "ContainerInherit, ObjectInherit","none", "Allow")
# Compile the rule into an ACL
$homeacl.AddAccessRule($rule)
# Apply the new rules to the home directory
Set-Acl $homedir $homeacl
# Check to make sure the permissions were applied, if not then repeat the loop
foreach($acl in $homeacl.Access)
{
    if($acl.IdentityReference.Value.substring(4) -eq $SAMAccountName)
    {
        $aclSet = $true
    }
}
# Start sleep to avoid script bashing
start-sleep -s 1
} until ($aclSet -eq $true)
# Revert error logging to normal state
$ErrorActionPreference = 'continue'
#----------------------------------------------------------------------------


#----------------------------------------------------------------------------
# 2. Force change password at next logon
#----------------------------------------------------------------------------
#Get the user that was created
$user = get-aduser $samaccountname
#Set the password on the account
Set-ADUser -Identity $user -ChangePasswordAtLogon $true
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# 3. Set the TsProfilePath
#----------------------------------------------------------------------------
#get the distinguished name of the user
$path = $user.DistinguishedName
#Create a profile path for LDAP for the user
$TsProfilePath = "LDAP://"+$path
#Create a tsprofile path for the user
$tshomedirpath = "PATH TO CITRIX PROFILE\"+$samaccountname
#Get the tsuser variable from ADSI
$tsuser = [ADSI] $TsProfilePath
#Set the TsProfilePath, TsHomeDrive, TsHomeDirectory
$tsuser.psbase.Invokeset("terminalservicesprofilepath",$ProfilePath)
$tsuser.psbase.Invokeset("TerminalServicesHomeDrive", "Y:")
$tsuser.psbase.Invokeset("TerminalServicesHomeDirectory", $tshomedirpath)
$tsuser.setinfo()
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
#3.1 Other Account Settings
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# Set the employee type in the attribute employeetype
switch($EmpType)
{
    "Temp" {Set-ADUser $user -replace @{employeeType="Temp"}}
    "Contractor" {Set-ADUser $user -replace @{employeeType="Contractor";EmployeeNumber="000"}}
    "Employee" {Set-ADUser $user -replace @{employeeType="Employee"}}
}
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# Get the manager's name and then get their AD account
$ReqSplit = $Requestor.split(" ")
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# Catch any user who doesn't have the normal First Last name
if($ReqSplit.length -eq 2)
{
    $ManagerFirst = $ReqSplit[0].trim()
    $ManagerLast = $ReqSplit[1].trim()
}
else
{
    $ManagerFirst = Read-Host "Enter Manager First Name"
    $ManagerLast = Read-Host "Enter Manager Last Name"
}
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# Get the manager's AD Account
if(Get-AdUser -filter {GivenName -eq $managerFirst -and Surname -eq $ManagerLast})
{
    # Get the manager AD account
    $managerAD = Get-ADUser -filter {GivenName -eq $managerfirst -and Surname -eq $managerlast} -properties DistinguishedName,EmailAddress,SamAccountName
    
    # Check to see if more than 1 person has that name
    if($managerAD.length -ge 2)
    {
        # If so and the server found exactly 2 accounts then check to see if one is administrative
        if($managerAD.length -eq 2 -and $managerAD[0].DistinguishedName -like "*Admin*" -or $managerAD[1].DistinguishedName -like "*admin*")
        {
            # Check to see which account is admin and assign the other account to the variable
            if($managerAD[0].DistinguishedName -like "*admin*")
            {
                $ManagerAD = Get-ADUser $ManagerAD[1] -properties DistinguishedName,EmailAddress,SamAccountName
            }
            if($managerAD[1].DistinguishedName -like "*admin*")
            {
                $ManagerAD = Get-ADUser $ManagerAD[0] -properties DistinguishedName,EmailAddress,SamAccountName
            }
        }
        # If two or more accounts were found and one does not appear to be admin then ask for username
        else
        {
            $ManagerSAM = Read-Host "Found more than one user matching that name, please specify username of manager"
            $managerAD = Get-ADUser $ManagerSAM -properties DistinguishedName,EmailAddress,SamAccountName
        }
    }
}
# If the account cannot be found at all (ex David instead of Dave was entered) Then ask for the account name
else
{
    $ManagerSAM = Read-Host "Cannot Find Manager in AD... Please enter username manually."
    $managerAD = Get-ADUser $ManagerSAM -properties DistinguishedName,EmailAddress,SamAccountName
}

# Set the manager in AD
Set-AdUser $user -replace @{manager=$managerAD.DistinguishedName}
#----------------------------------------------------------------------------


#----------------------------------------------------------------------------
# 4. Enable Account
Enable-ADAccount -Identity $user
#----------------------------------------------------------------------------

#============================================================================

#Add a wait timer
do
{
Start-Sleep -s 1
$enableCheck = Get-ADUser $user -properties Enabled
}
until($enableCheck.Enabled -eq $true)

#============================================================================
# EXCHANGE/CREATE NEW MAILBOX FOR USER
#============================================================================
# Steps
# 1. Create connection to exchange
# 2. Detemine mailbox then enable
# 3. Disconnect from exchange server
Write-Host "Stepping into Exchange section"
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# 1. Create connection to exchange
#----------------------------------------------------------------------------
#Write-Host "Connecting to exchange server"
#Create a connection to the Exchange server in order to enable the new user's email account
#$ExchSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "PATH TO EXCHANGE SERVER"
#Open up the connection
#Import-PSSession $ExchSession
#Add a wait timer for connecting properly
#Start-Sleep -s 5
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# 2. Detemine mailbox then enable
#----------------------------------------------------------------------------
#Enable the mailbox of the new user
$ErrorActionPreference = 'silentlyContinue'
do
{
Start-Sleep -s 2
if($template -eq "Corp" -or $template -eq "West" -or $template -eq "South")
{
    switch($row % 4)
    {
        0 {Enable-Mailbox -Identity $usrPName -Alias $samaccountname -Database 'Database1'}
        1 {Enable-Mailbox -Identity $usrPName -Alias $samaccountname -Database 'Database2'}
        2 {Enable-Mailbox -Identity $usrPName -Alias $samaccountname -Database 'Database3'}
        3 {Enable-Mailbox -Identity $usrPName -Alias $samaccountname -Database 'Database4'}
    }
}
if($template -eq "East" -or $template -eq "French")
{
    switch($row % 2)
    {
        0 {Enable-Mailbox -Identity $usrPName -Alias $samaccountname -Database 'Database5'}
        1 {Enable-Mailbox -Identity $usrPName -Alias $samaccountname -Database 'Database6'}
    }
}
}
until (Get-Mailbox -Identity $samaccountname)
$ErrorActionPreference = 'continue'


$sync = get-CASMailbox -Identity $samaccountname
if($sync.ActiveSyncEnabled -eq $true)
{
    Set-CASMailbox -Identity $samaccountname -ActiveSyncEnabled $false
}

#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# 3. Disconnect from the exchange server
#Remove-PSSession $ExchSession
#----------------------------------------------------------------------------
Write-Host "Exchange portion complete"
# END EXCHANGE PORTION OF SCRIPT
#============================================================================



#============================================================================
# ASSIGNING PERMISSIONS TO THE NEW USER
#============================================================================
# Steps
# 1. Get template based on location
# 2. Assign template permissions to user
# 3. Run mirror User script.
Write-Host "Stepping into the permissions section"
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# 1. Get template based on location
#----------------------------------------------------------------------------


#----------------------------------------------------------------------------
# 2. Assign template permissions to user
#----------------------------------------------------------------------------
Write-Host "Assigning Template permissions"
#Get all memberships that the appropriate Template has and add them
$source = Get-AdUser $templateUser -Properties memberof
#step through all memberships of template and apply them
foreach($group in $source.memberof)
{
    #Add template groups to the new user
    add-adgroupMember $group -members $user
}
#----------------------------------------------------------------------------

#----------------------------------------------------------------------------
# 3. Run mirror user script
#----------------------------------------------------------------------------
if($mirrorUser)
{
    Write-Host "Found user to mirror, running Memberships.ps1"
    .\Functions\Memberships.ps1
}
#----------------------------------------------------------------------------
#END PERMISSIONS SECTION OF SCRIPT
#============================================================================

# AUTOMATED EMAIL SEND-------------------------------------------------------
$ButtonLabel = "Is this a rehire?"
$Rehire = "N"
$Rehire = . $anotherUserButton

$ButtonLabel = "Do you want to send an automated new hire email?"
$SendEmail = "N"
$SendEmail = . $anotherUserButton
if($SendEmail -eq "Y")
{
    . $autoEmail
}
#----------------------------------------------------------------------------

$UserCreatedList += $user.UserPrincipalName + ";" + $lyncloc

$ButtonLabel = "Do you want to create another user?"
$AnotherUser = . $anotherUserButton
#$AnotherUser = Read-Host "Another User to create? Please enter Y or N:"
}

while($AnotherUser -eq "Y")


#Properly close and exit the excel process
$workbook.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
Remove-Variable ws
Remove-Variable workbook
Remove-Variable Excel

#Remove the Exchange Session
Remove-PSSession $ExchSession

}