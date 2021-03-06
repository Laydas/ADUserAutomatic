######################################################################
# NewHireEmail.ps1
# Written by Iain Letourneau
# Date last edited March 30 2016
# Contact: Iain.Letourneau@gmail.com
# 
# This script will take input from the CreateUser.ps1 and compile the
# standard new hire email and send to the correct people based on
# where the new account's office location is.
# If the employee was set to contractor then HR is included in the email
######################################################################

#---------------------------------------------------------------------
# Import the data from the VariableChange.ps1 file
. ..\VariablesChange.ps1

# Get the PrimaryITAdmin email
$adPrimary = Get-ADUser $PrimaryITAdmin -properties EmailAddress
$PrimaryString = $adPrimary.SAMAccountName + " <" + $adPrimary.EmailAddress + ">"

# Get the SAP Team emails
$SAPArray = @()
foreach($user in $SAPNewHireTeam)
{
    $aduser = Get-ADUser $user -properties EmailAddress
    $SAPArray += $aduser.SAMAccountName + " <" + $aduser.EmailAddress + ">"
}
# Get the Ontario Team emails
$OntarioArray = @()
foreach($user in $OntarioNewHireTeam)
{
    $aduser = Get-ADUser $user -properties EmailAddress
    $OntarioArray += $aduser.SAMAccountName + " <" + $aduser.EmailAddress + ">"
}
# Get the Quebec Team emails
$QuebecArray = @()
foreach($user in $QuebecNewHireTeam)
{
    $aduser = Get-ADUser $user -properties EmailAddress
    $QuebecArray += $aduser.SAMAccountName + " <" + $aduser.EmailAddress + ">"
}
# Get the Western Canada Team emails
$WestCanArray = @()
foreach($user in $WestCanNewHireTeam)
{
    $aduser = Get-ADUser $user -properties EmailAddress
    $WestCanArray += $aduser.SAMAccountName + " <" + $aduser.EmailAddress + ">"
}
# Get the Southern Team emails
$SouthArray = @()
foreach($user in $SouthNewHireTeam)
{
    $aduser = Get-ADUser $user -properties EmailAddress
    $SouthArray += $aduser.SAMAccountName + " <" + $aduser.EmailAddress + ">"
}
# Get the Atlantic Team emails
$AtlanticArray = @()
foreach($user in $AtlanticNewHireTeam)
{
    $aduser = Get-ADUser $user -properties EmailAddress
    $AtlanticArray += $aduser.SAMAccountName + " <" + $aduser.EmailAddress + ">"
}
#---------------------------------------------------------------------

#---------------------------------------------------------------------
# Get the manager's email string compiled

# Create the manager email string for sending mail
$to = $managerAD.SamAccountName + " <" + $managerAD.emailaddress + ">"
#---------------------------------------------------------------------

$newlogo = "PATH TO LOGO FOR EMAIL"

#---------------------------------------------------------------------

#Compile the CC recipients--------------------------------------------

[string[]]$cc = $SAPArray
[string[]]$bcc = $PrimaryString

if($empType -ne "Contractor")
{
    switch($EmailLoc)
    {
        "Ontario" {foreach($TeamEmail in $OntarioArray){[string[]]$cc += $TeamEmail}}
        "Quebec" {foreach($TeamEmail in $QuebecArray){[string[]]$cc += $TeamEmail}}
        "WestCan" {foreach($TeamEmail in $WestCanArray){[string[]]$cc += $TeamEmail}} 
        "South" {foreach($TeamEmail in $WestCanArray){[string[]]$cc += $TeamEmail}; foreach($TeamEmail in $SouthArray){[string[]]$cc += $TeamEmail}}
        "Atlantic" {foreach($TeamEmail in $AtlanticArray){[string[]]$cc += $TeamEmail}}
    }
    [string[]]$bcc += "eLearn <elearn@example.com>"
}
#---------------------------------------------------------------------


#---------------------------------------------------------------------
# Compile the bcc and from--------------------------------------------

$from = "security <security@example.com>"
#---------------------------------------------------------------------


#---------------------------------------------------------------------
# Enter the server smtp address---------------------------------------
$smtp = "smtp.example.com"
#---------------------------------------------------------------------


#---------------------------------------------------------------------
# Compile the subject string------------------------------------------
if($Rehire -eq "N")
{
    $subject = "Ticket #" + $ticket + " - New Hire - " + $First + " " + $Last + " - Windows User Name and Password"
}
else
{
    $subject = "Ticket #" + $ticket + " - Rehire - " + $First + " " + $Last + " - Windows User Name and Password"
}
#---------------------------------------------------------------------


#---------------------------------------------------------------------
# Compile the email message to look exactly like the templates--------
$message = @"
<head>
<style>
p.MsoNormal, li.MsoNormal, div.MsoNormal {margin:0in; margin-bottom:.0001pt; font-size:11.0pt; font-family:"Calibri","sans-serif";}
a:link, span.MsoHyperlink {mso-style-priority:99; color:blue; text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed {mso-style-priority:99; color:purple; text-decoration:underline;}
p.MsoAcetate, li.MsoAcetate, div.MsoAcetate {mso-style-priority:99; mso-style-link:"Balloon Text Char"; margin:0in; margin-bottom:.0001pt; font-size:11.0pt; font-family:"Calibri","sans-serif";}
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph {mso-style-priority:34; margin-top:0in; margin-right:0in; margin-bottom:0in; margin-left:.5in; margin-bottom:.0001pt; font-size:11.0pt; font-family:"Calibri","sans-serif";}
span.BalloonTextChar {mso-style-name:"Balloon Text Char"; mso-style-priority:99; mso-style-link:"Balloon Text"; font-family:"Tahoma","sans-serif";}
span.TextedebullesCar {mso-style-name:"Texte de bulles Car"; mso-style-priority:99; mso-style-link:"Texte de bulles"; font-family:"Tahoma","sans-serif";}
p.Textedebulles, li.Textedebulles, div.Textedebulles {mso-style-name:"Texte de bulles"; mso-style-link:"Texte de bulles Car"; margin:0in; margin-bottom:.0001pt; font-size:11.0pt; font-family:"Calibri","sans-serif";}
span.EmailStyle22 {mso-style-type:personal-compose; font-family:"Calibri","sans-serif"; color:#365F91;}
span.EmailStyle23 {mso-style-type:personal; font-family:"Calibri","sans-serif"; color:#365F91;}
.MsoChpDefault {mso-style-type:export-only; font-size:10.0pt; font-family:"Calibri","sans-serif";}
</style>
</head>
<body lang=EN-US link=blue vlink=purple>
<p class=MsoNormal>
<span style='color:#365F91'>
"@
$message += "Hello " + $ManagerFirst
$message += @"
, <o:p></o:p></span></p><p class=MsoNormal><span style='color:#365F91'><o:p>&nbsp;</o:p></span></p>
<p class=MsoNormal><span lang=EN-CA style='color:#365F91'>Please provide the following log on user credentials to your new hire.<o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-CA style='color:#365F91'><o:p>&nbsp;</o:p></span></p>
<p class=MsoNormal><b><span lang=EN-CA style='color:#365F91'>Windows log on credentials: <o:p></o:p></span></b></p>
<p class=MsoNormal><span lang=EN-CA style='color:#365F91'>
"@
$message += "User name: " + $SAMAccountName
$message += @"
</span>
<span lang=EN-CA style='font-family:"Arial","sans-serif"'><o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-CA style='color:#365F91'>
"@
$message += "Password: " + $Password
$message += @"
<o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-CA style='color:#365F91'>
"@
$message += "Email: " + $Email
$message += @"
</span>
<span style='font-size:10.0pt;font-family:`"Arial`",`"sans-serif`"'><o:p></o:p></span></p>
<p class=MsoNormal><span lang=EN-CA style='color:#365F91'><o:p>&nbsp;</o:p></span></p>
<p class=MsoNormal><span lang=EN-CA style='color:#365F91'><o:p>&nbsp;</o:p></span></p>
<p class=MsoNormal>
<span lang=EN-CA style='color:#365F91'>Should you have any questions, please create a Footprints support ticket at </span>
<span lang=FR-CA style='color:#365F91'><a href="http://service.example.com">
<span lang=EN-CA>http://service.example.com</span></a></span>
<span lang=EN-CA style='color:#365F91'> or call<o:p></o:p></span></p>
<p class=MsoNormal><span lang=FR-CA style='color:#365F91'>Toll-free:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 1-866-349-3247<o:p></o:p></span></p>
<p class=MsoNormal><span lang=FR-CA style='color:#365F91'>Local:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 403-806-7344<o:p></o:p></span></p>
<p class=MsoNormal><img border=0 width=83 height=77 id="ExampleLogo" src=
"@
$message += "`"" + $newlogo + "`""
$message += @"
alt=""><o:p></o:p></p>
<p class=MsoNormal><a href="http://www.example.com/">www.example.com</a><o:p></o:p></p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<p class=MsoNormal><span style='color:#365F91'>To access the Example Web Portal, click</span> <a href="https://start.example.com/">here</a>.<o:p></o:p></p><p class=MsoNormal><o:p>&nbsp;</o:p>
</p>
</body>
"@
# Finish Compiling Email Message--------------------------------------


#---------------------------------------------------------------------
#Send the compiled message!!------------------------------------------
send-mailmessage -to $to -Cc $cc -Bcc $bcc -from $from -subject $subject -BodyAsHtml $message -smtpServer $smtp
#---------------------------------------------------------------------