<<<<<<< HEAD
<#
    .Synopsis
        Create User Account and Mailbox in Exchange
    .Description
        Copies User Account if specified and creates mailbox and sets attributes such as manager, ProxyAddress etc
     .Example
#>    

#Import AD and Exchange
Import-Module ActiveDirectory
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.domain.com.au/PowerShell/ -Authentication Kerberos 
Import-PSSession $Session -AllowClobber

#Import VB
Add-Type -AssemblyName Microsoft.VisualBasic
$vb = [Microsoft.VisualBasic.Interaction]

### Variables ###
$samaccount_to_disable = $vb::inputbox("Enter SAMAccount Name to Disable")
$DisplayName = ($samaccount_to_disable).Name
$Ticket_Number = $vb::inputbox("Enter Ticket Number")
$logpath = "\\server\c$\Temp\DisabledUsersTest\ "
$logfile = $logpath + "\$samaccount_to_disable.txt"
$datestamp = ((Get-Date).ToString('dd-MM-yyyy'))

### Check if user exists ###

$User = $(try {Get-ADUser $samaccount_to_disable -Properties Name, EmailAddress, Manager | Select-Object Name} catch {$null})

If ($User -eq $Null) {
    Write-Host "User doesn't Exist in AD, Please run script again"
}
Else {
    Write-Host "User found in AD, Continuing"
}

### Get Manager ###

$Manager = $(try {(Get-ADUser (Get-ADUser $samaccount_to_disable -Properties manager).manager).SamAccountName} catch {$null})
$ManagerEmail = $Manager.mail
If ($Manager -eq $Null) {
    Write-Host "No Manager set,"
    $samaccount_to_forward_email = $vb::inputbox("Enter SAMAccount Name to forward email to")
}
Else {
    Write-Host "Manager set, Continuing"
}

### Disable User

Disable-ADAccount $samaccount_to_disable
Write-Host -ForegroundColor Yellow "Disabling Account..."

### Backup group memberships to text file ###

"Disabled by Name " + $datestamp + " $Ticket_Number " | Out-File $logfile -append
$Group_MembershipNames = $Group_Memberships = Get-ADPrincipalGroupMembership -Identity $samaccount_to_disable | Select-Object Name | Out-File $logfile -Append
$Group_Memberships = Get-ADPrincipalGroupMembership -Identity $samaccount_to_disable  | Where-Object { $_.Name -notcontains "Domain Users" }
If ($Group_Memberships -ne $null) {
    foreach ($group in $Group_Memberships) {
        # Remove each group membership from the user
        Write-Host -ForegroundColor Yellow    "Removing user from $($group.name) "
        $Group_Memberships | Remove-ADGroupMember -Members $samaccount_to_disable
        Write-Host -ForegroundColor Green "$($group.name) Removed"
    }
}
Else {
    Write-Host -ForegroundColor Red "Group Membership still exist please remove manually"
}



### Get Current OU and split it to the site name. ###

$CurrentOU = ($samaccount_to_disable).distinguishedName.Split(',')[3].substring(0)

### Disabled OU ###
$DisabledOU = "OU=Disabled Users, DC=domain,Dc=com,DC=au"

### If multiple Disabled Users OUs as per AD Structure ###

$Site1OU = "OU=Disabled Users,OU=Users,OU=Site1,DC=domain,DC=com,DC=au"
$Site2OU = "OU=Disabled Users,OU=Site2,DC=domain,DC=com,DC=au"
$Site3OU = "OU=Disabled Users,OU=Site3,DC=domain,DC=com,DC=au"
$Site4OU = "OU=Disabled Users,OU=Site4,DC=domain,DC=com,DC=au"

#switch -Wildcard ($CurrentOU) { 
#    "*Site1*" { $DisableOU = $Site1OU; break }
#    "*Site2*" { $DisableOU = $Site2OU; break }
#    "*Site3*" { $DisableOU = $Site3OU; break }
#    "*Site4*" { $DisableOU = $Site4OU; break } 
#    default { Write-Host "No Path specified" }
#} 
#$DisabledOU

### Determine OU user is in and which Disabled Users OU to move them to ##

If ($CurrentOU -like "*OU=Site1*") {
    $DisabledOU = $Site1OU
    $DisabledOU
}
Elseif ($CurrentOU -like "*OU=Site2*") {
    $DisabledOU = $Site2OU
    $DisabledOU
}
Elseif ($CurrentOU -like "*OU=Site3*") {
    $DisabledOU = $Site3OU
    $DisabledOU
}
Elseif ($CurrentOU -like "*OU=Site4*") {
    $DisabledOU = $Site4OU
    $DisabledOU
}
Else {
    "Path Not Found"
}

### Move User to Disabled Users OU ###
$samaccount_to_disable | Move-ADObject -TargetPath $DisabledOU

### Hide from Global Address List
Set-Mailbox -Identity $samaccount_to_disable -HiddenFromAddressListsEnabled $true

### Forward Email ###
If ($Manager -eq $Null) {
    Set-Mailbox -Identity "$DisplayName" -ForwardingSMTPAddress "$mailbox_to_forward_email"
}
Else {
    Set-Mailbox Identity "$DisplayName" -ForwardingSMTPAddress $ManagerEmail
=======
#==========================================================================
#
# NAME: DisableAccount.ps1
#
# AUTHOR: mcaulr09
# 
#
# SUMMARY:
#
# Powershell Script to Disable an account as per Remove User standard process.
#
# DESCRIPTION:
#
# Powershell Script to Disable an account as per Remove User standard process.
# This is done by disasbling the account, moving it to the disabled users OU
# Removing group memberships, updating description, hiding mailbox from exchange
# and forward email to manager/specified contact.

# VERSION HISTORY:
# 1.0 4/01/2018 - Rachel McAuliffe - Work Commenced

#Import AD and Exchange
Import-Module ActiveDirectory
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.domain.com.au/PowerShell/ -Authentication Kerberos 
Import-PSSession $Session -AllowClobber


#Import VB
Add-Type -AssemblyName Microsoft.VisualBasic
$vb = [Microsoft.VisualBasic.Interaction]

### Variables ###
$samaccount_to_disable = $vb::inputbox("Enter SAMAccount Name to Disable")
$DisplayName = (Get-ADUser $samaccount_to_disable).Name
$Ticket_Number = $vb::inputbox("Enter Ticket Number")
$Logfiletime = (Get-Date).ToString('dd-MM-yyyy')
$logpath = "\\server\c$\Temp\DisabledUsersTest\ "
$logfile = $logpath + "\$samaccount_to_disable.txt"
$datestamp = ((Get-Date).ToString('dd-MM-yyyy'))

### Check if user exists ###

$User = $(try {Get-ADUser $samaccount_to_disable -Properties Name,EmailAddress,Manager | Select Name} catch {$null})

If ($User -eq $Null)
{
   Write-Host "User doesn't Exist in AD, Please run script again"
}
Else
{
   Write-Host "User found in AD, Continuing"
}

### Get Manager ###

$Manager = $(try {(Get-ADUser (Get-ADUser $samaccount_to_disable -Properties manager).manager).SamAccountName} catch {$null})
$ManagerEmail = $Manager.mail
If ($Manager -eq $Null)
{
    Write-Host "No Manager set,"
    $samaccount_to_forward_email = $vb::inputbox("Enter SAMAccount Name to forward email to")
}
Else
{
    Write-Host "Manager set, Continuing"
}

### Disable User

Disable-ADAccount $samaccount_to_disable
Write-Host -ForegroundColor Yellow "Disabling Account..."

### Backup group memberships to text file ###

"Disabled by Name " + $datestamp + " $Ticket_Number " | Out-File $logfile -append
$Group_MembershipNames = $Group_Memberships = Get-ADPrincipalGroupMembership -Identity $samaccount_to_disable | Select Name | Out-File $logfile -Append
$Group_Memberships = Get-ADPrincipalGroupMembership -Identity $samaccount_to_disable  | Where-Object { $_.Name -notcontains "Domain Users" }
    If ($Group_Memberships -ne $null)
    {
    foreach($group in $Group_Memberships){
    # Remove each group membership from the user
    Write-Host -ForegroundColor Yellow    "Removing user from $($group.name) "
    $Group_Memberships | Remove-ADGroupMember -Members $samaccount_to_disable
    Write-Host -ForegroundColor Green "$($group.name) Removed"
        }
    }
   Else
    {
    Write-Host -ForegroundColor Red "Group Membership still exist please remove manually"
}



### Get Current OU and split it to the site name. ###

$CurrentOU = (Get-AdUser $samaccount_to_disable).distinguishedName.Split(',')[3].substring(0)

### Disabled OU ###
$DisabledOU = "OU=Disabled Users, DC=domain,Dc=com,DC=au"

### If multiple Disabled Users OUs as per AD Structure ###

$Site1OU = "OU=Disabled Users,OU=Users,OU=Site1,DC=domain,DC=com,DC=au"
$Site2OU = "OU=Disabled Users,OU=Site2,DC=domain,DC=com,DC=au"
$Site3OU = "OU=Disabled Users,OU=Site3,DC=domain,DC=com,DC=au"
$Site4OU = "OU=Disabled Users,OU=Site4,DC=domain,DC=com,DC=au"

#switch -Wildcard ($CurrentOU) { 
#    "*Site1*" { $DisableOU = $Site1OU; break }
#    "*Site2*" { $DisableOU = $Site2OU; break }
#    "*Site3*" { $DisableOU = $Site3OU; break }
#    "*Site4*" { $DisableOU = $Site4OU; break } 
#    default { Write-Host "No Path specified" }
#} 
#$DisabledOU

### Determine OU user is in and which Disabled Users OU to move them to ##

If($CurrentOU -like "*OU=Site1*")
{
$DisabledOU = $Site1OU
$DisabledOU
}
Elseif($CurrentOU -like "*OU=Site2*")
{
$DisabledOU = $Site2OU
$DisabledOU
}
Elseif($CurrentOU -like "*OU=Site3*")
{
$DisabledOU = $Site3OU
$DisabledOU
}
Elseif($CurrentOU -like "*OU=Site4*")
{
$DisabledOU = $Site4OU
$DisabledOU
}
Else
{
    "Path Not Found"
}

### Move User to Disabled Users OU ###
Get-ADUser $samaccount_to_disable | Move-ADObject -TargetPath $DisabledOU

### Hide from Global Address List
Set-Mailbox -Identity $samaccount_to_disable -HiddenFromAddressListsEnabled $true

### Forward Email ###
If ($Manager -eq $Null)
{
Set-Mailbox -Identity "$DisplayName" -ForwardingSMTPAddress "$mailbox_to_forward_email"
}
Else
{
Set-Mailbox Identity "$DisplayName" -ForwardingSMTPAddress $ManagerEmail
>>>>>>> 9543955e5800114e65173d066837a34bddd57161
}   