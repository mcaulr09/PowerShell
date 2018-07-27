[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") | Out-null
$vb = [Microsoft.VisualBasic.Interaction]
$DebugPreference = 'Inquire'
Import-Module ActiveDirectory$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.mcauliffe-systems.com/PowerShell/ -Authentication Kerberos 
Import-PSSession $Session -AllowClobber
##### Variables #####

$SharedMailboxName = $vb::inputbox("Shared Mailbox Name")
$FirstName = ($SharedMailboxName.split(" ")[0])
$LastName = ($SharedMailboxName.Substring($SharedMailboxName.IndexOf(" ") +1))
$Alias = ($SharedMailboxName -replace "\s","")
$UserPrincipalName = ($SharedMailboxName + "@mcauliffe-systems.com"-replace "\s","")
$OU = ('OU=Test,DC=mcauliffe-systems,DC=com')
$FAGroupName = ("MBS - " + $SharedMailboxName + " - FA")
$SAGroupName = ("MBS - " + $SharedMailboxName + " - SA")
$GroupsOU = ('OU=Test,DC=mcauliffe-systems,DC=com')
$FADescription = ("Full Access for " + $SharedMailbox)
$SADescription = ("Send As for " + $SharedMailbox)
$Password = ConvertTo-SecureString -string “Password” -asPlainText -Force


if ($SharedMailboxName -notmatch "\s")
{
Write-Host "Please space out shared mailbox name"
Exit
}
Else
{
{continue}
} 

##### Generate Random Password from DinoPass # Credit to Chris Spencer

#$web = New-Object Net.WebClient #Generates powershell web client
#$WebProxy = New-Object System.Net.WebProxy("http://DMZIS01.ext.iwf.com.au:8080",$true)
#$webproxy.UseDefaultCredentials = $true
#$web.proxy = $webproxy
#$web.Headers.Add("Cache-Control", "no-cache");
#$PwdString= $web.DownloadString("http://www.dinopass.com/password/simple")
#$PwdString = $PwdString.substring(0,1).toUpper() + $PwdString.substring(1)
#$Password = ConvertTo-SecureString -String $PwdString -AsPlainText -Force 
##### Check if Name, Mailbox and Groups already exist #####$DispName = $(try {Get-ADUser -Filter{displayName -like $SharedMailboxName} -Properties SamAccountName} catch {$null})

if ($DispName -eq $null)
{
    Write-Host "Please enter Desired Mailbox Name"  
}
Else
{
    Write-Host "Creating Mailbox"
}$Mailbox = $(try {Get-ADUser $Alias -Properties * | Select Name} catch {$null}) If ($Mailbox -eq $Null){   Write-Host "Mailbox doesn't exist"}Else{   Write-Host "Mailbox found"} 

$Groups = @("$FAGroupName","$SAGroupName")
$(try {Get-ADUser $Groups -Properties | select Name} catch {$null}) 


If ($Groups -eq $Null)
{
    Write-Host "No Groups Found"
}
Else
{
    Write-Host "Groups Already exist"
}
  
##### Create Mailbox #####
New-Mailbox -Name $SharedMailboxName -DisplayName $SharedMailboxName -FirstName $Firstname -LastName $LastName -Alias $Alias -UserPrincipalName $UserPrincipalName -OrganizationalUnit $OU -Password $Password -ResetPasswordOnNextLogon $false -Verbose

##### Confirm Mailbox Created #####
Get-Mailbox -Identity $Alias

##### Create Groups #####
New-ADGroup $FAGroupName -Path $GroupsOU -GroupScope Global -GroupCategory Security -Description $FADescription -Verbose
New-ADGroup $SAGroupName -Path $GroupsOU -GroupScope Global -GroupCategory Security -Description $SADescription -Verbose


##### Confirm Groups Created
Get-ADGroup $FAGroupName
Get-ADGroup $SAGroupName

##### Apply Permissions #####
Add-MailboxPermission –Identity $SharedMailboxName –user "MCAULIFFE\$FAGroupName" –AccessRights "FullAccess" -Verbose
Add-ADPermission –Identity $SharedMailboxName –user "MCAULIFFE\$SAGroupName" –ExtendedRights "Send As" -Verbose

##### Verify Permissions #####
Get-MailboxPermission –Identity $UserPrincipalName