############################################################
# VariablesChange.ps1
# Written by Iain Letourneau
# Date last edited March 30 2016
# Contact: Iain.Letourneau@gmail.com
#
# This is a text file that contains all the variables that 
# might need to be changed for any of the scripts
#
# The following is the required convention for variables
# $VariableName = "samaccountname"
# or 
# $VariableArray = @(
#   "one", 
#   "two", 
#   "three" <-notice there is no comma(,) on the last entry
# )
#
############################################################

# This is for the individual performing most of the newhire/termination tasks
$primaryITAdmin = "sskelton"
$secondaryITAdmin = "jsallay"

#============================================================
#**************** NEW HIRE EMAIL SECTION ********************
#============================================================
# List of team members for region specific hiring process

# WestCan
$WestCanNewHireTeam = @(
    "example1",
    "example2",
    "example3",
    "example4"
)

# Ontario
$OntarioNewHireTeam = @(
    "example1"
)

# Quebec
$QuebecNewHireTeam = @(
    "example1",
    "example2"
)

# South
$SouthNewHireTeam = @(
    "example1"
)

#Atlantic
$AtlanticNewHireTeam = @(
    "example1"
)

# SAPNewHireTeam - Team responsible for setting up the SAP side of new hires
$SAPNewHireTeam = @(
    "example1",
    "example2"
)
#============================================================