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
$Ticket_Number = $vb::inputbox("Enter Ticket Number")
$logpath = "\\server\c$\Temp\DisabledUsersTest\ "
$logfile = $logpath + "\$samaccount_to_disable.txt"
$datestamp = ((Get-Date).ToString('dd-MM-yyyy'))

### Check if user exists ###

$User = $(try {Get-ADUser $samaccount_to_disable -Properties SamAccountName,Name,distinguishenName,EmailAddress,Manager} catch {$null})

If ($User -eq $Null) {
    Write-Host -ForegroundColor Red "User to copy doesn't Exist in AD, Please run script again"
}
Else {
    Write-Host -ForegroundColor Green "User to copy found in AD, Continuing"
}

### Get Manager ###

$Manager = $(try {(Get-ADUser (Get-ADUser $samaccount_to_disable -Properties manager).manager)} catch {$null})
$ManagerName = $Manager.Name
$ManagerAccount = $Manager.SamaccountName
If ($Manager -eq $Null) {
    Write-Host -ForegroundColor Red "No Manager set,"
    $samaccount_to_forward_email = $vb::inputbox("Enter SAMAccount Name to forward email to")
}
Else {
    Write-Host -ForegoundColor Green "Manager set, Continuing"
}

### Disable User
Disable-ADAccount $samaccount_to_disable
Write-Host -ForegroundColor Yellow "Disabling Account..."

### Set Description ###
Set-ADUser $samaccount_to_disable -Description "Disabled by - $datestamp - $Ticket_Number"

### Backup group memberships to text file ###
"Disabled by  " + $datestamp + " $Ticket_Number " | Out-File $logfile -append
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

#$CurrentOU = ($User).distinguishedName.Split(',')[3].substring(0)

### Disabled OU ###
$DisabledOU = "OU=Disabled Users,DC=domain,Dc=com,DC=au"

### If multiple Disabled Users OUs as per AD Structure ###

#$Site1OU = "OU=Disabled Users,OU=Users,OU=Site1,DC=domain,DC=com,DC=au"
#$Site2OU = "OU=Disabled Users,OU=Site2,DC=domain,DC=com,DC=au"
#$Site3OU = "OU=Disabled Users,OU=Site3,DC=domain,DC=com,DC=au"
#$Site4OU = "OU=Disabled Users,OU=Site4,DC=domain,DC=com,DC=au"

#switch -Wildcard ($CurrentOU) { 
#    "*Site1*" { $DisableOU = $Site1OU; break }
#    "*Site2*" { $DisableOU = $Site2OU; break }
#    "*Site3*" { $DisableOU = $Site3OU; break }
#    "*Site4*" { $DisableOU = $Site4OU; break } 
#    default { Write-Host "No Path specified" }
#} 
#$DisabledOU

### Determine OU user is in and which Disabled Users OU to move them to ##

#If ($CurrentOU -like "*OU=Site1*") {
#    $DisabledOU = $Site1OU
#    $DisabledOU
#}
#Elseif ($CurrentOU -like "*OU=Site2*") {
#    $DisabledOU = $Site2OU
#    $DisabledOU
#}
#Elseif ($CurrentOU -like "*OU=Site3*") {
#    $DisabledOU = $Site3OU
#    $DisabledOU
#}
#Elseif ($CurrentOU -like "*OU=Site4*") {
#    $DisabledOU = $Site4OU
#    $DisabledOU
#}
#Else {
#    Write-Host -ForegroundColor Red "Path Not Found"
#}

### Move User to Disabled Users OU ###
Write-Host -ForegroundColor Yellow "Moving to Disabled Users OU"
$User | Move-ADObject -TargetPath $DisabledOU

### Hide from Global Address List
Write-Host -ForegroundColor Yellow "Hiding from Global Addres List"
Set-Mailbox -Identity $samaccount_to_disable -HiddenFromAddressListsEnabled $true

### Forward Email ###
If ($Manager -ne $Null) {
    $ManagerEmail = Get-ADUser $Manager -Properties Mail | Select Mail
    Set-Mailbox -Identity "$DisplayName" -ForwardingSMTPAddress "$ManagerEmail"    
}
Elseif {
    Set-Mailbox -Identity "$User" -ForwardingSMTPAddress "$mailbox_to_forward_email"
    }
Else
{
Write-Host ForgroundColor Red "Emails not forwarded please contact Manager"    
}

##### Send Account Details #####


#$smtp = "mail.domain.com.au"

#$to = "$ManagerName <$ManagerAccount@domain.com.au>" 

#$From = "Display Name <Name.Name@domain.com.au>"

#$Cc = "John Bolton <John.Bolton@domain.com.au>"

#$Bcc = "Rachel McAuliffe <Rachel.McAuliffe@domain.com.au>"

#$subject = " Re: $TicketNumber - $DisplayName Account Disabled"  
 
#$body = "Dear <b><font color=red>$ManagerName</b></font> <br><br>" 

#$body += "Kind Regards, <br>"
#$body += "Name Name <br><br>"

#### Now send the email using Send-MailMessage  
 
#send-MailMessage -SmtpServer $smtp -To $to -Bcc $Bcc -From $from -Subject $subject -Body $body -BodyAsHtml -Priority High 
