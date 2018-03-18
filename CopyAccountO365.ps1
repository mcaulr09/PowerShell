<##>
    .Synopsis
        Create User Account and Mailbox in Exchange
    .Description
        Copies User Account if specified and creates mailbox and sets attributes such as manager, ProxyAddress etc
     .Example
#>    

#Import AD
Import-Module ActiveDirectory

#Import VB
Add-Type -AssemblyName Microsoft.VisualBasic
$vb = [Microsoft.VisualBasic.Interaction]

#Variables
$samaccount_to_copy = $vb::inputbox("Enter SAMAccount Name to Copy")
$manageraccount = $vb::inputbox("Enter Manager SAMAccount Name")
$new_displayname = $vb::inputbox("Enter Display Name of new user")
$new_firstname = ($new_displayname.split(" ")[0])
$new_lastname = ($new_displayname.Substring($new_displayname.IndexOf(" ") + 1))
$new_name = $new_displayname
$CopyPath = $(try {(Get-AdUser $samaccount_to_copy).distinguishedName.Split(',', 2)[1]} catch {$null})
$ManagerPath = $(try {(Get-AdUser $manageraccount).distinguishedName.Split(',', 2)[1]} catch {$null})
$enable_user_after_creation = $true
$password_never_expires = $false
$cannot_change_password = $false
$ad_account_to_copy = $(try {Get-ADUser $samaccount_to_copy -Properties Description, Office, OfficePhone, StreetAddress, City, State, PostalCode, Country, Title, Company, Department, Manager, EmployeeID} catch {$null})
$ad_account_manager = $(try {Get-ADUser $manageraccount -Properties Office, OfficePhone, StreetAddress, City, State, PostalCode, Country, Company, Department} catch {$null})

##### Generate Random Password from DinoPass
$web = New-Object Net.WebClient #Generates powershell web client
$web.Headers.Add("Cache-Control", "no-cache");
$PwdString = $web.DownloadString("http://www.dinopass.com/password/simple")
$PwdString = $PwdString.substring(0, 1).toUpper() + $PwdString.substring(1)
$Password = ConvertTo-SecureString -String $PwdString -AsPlainText -Force  

### If not using DinoPass ###
#Function random-password ($length = 8) {
#    $punc = 46..46
#    $digits = 48..57
#    $letters = 65..90 + 97..122
#
    # Thanks to
    # https://blogs.technet.com/b/heyscriptingguy/archive/2012/01/07/use-pow
#    $password = get-random -count $length `
#        -input ($punc + $digits + $letters) |
#        % -begin { $aa = $null } `
#        -process {$aa += [char]$_} `
#        -end {$aa}
#
#    return $password
#}

#$Password = random-password

#####Check accounts exist#####
$User = $(try {Get-ADUser $samaccount_to_copy -Properties * | Select-Object Name} catch {$null})
 
If ($User -eq $Null) {
    Write-Host "User doesn't Exist in AD"
}
Else {
    Write-Host "User found in AD"
}
 
$Manager = $(try {Get-ADUser $manageraccount -Properties * | Select-Object Name} catch {$null})
 
If ($Manager -eq $Null) {
    "Manager doesn't Exist in AD"
}
Else {
    "Manager found in AD"
}

##### Generate Username - FirstInitialLastName
#function Generate-username($new_displayname)
#{
#    $count = 0
#
#    do
#   {
#        $count++
#        $new_samaccountname = "$($new_firstname.substring(0,$count))$new_lastname"
#    } until ( ! (Get-aduser -Filter {SamAccountName -eq $new_samaccountname} ) )
#
#    $new_samaccountname.ToLower()
#} 

##### Generate Username - FirstNameFirstInitialLastName #####
#function Generate-username($new_displayname)
#{
#    $count = 0
#
#    do
#    {
#        $count++
#        $new_samaccountname = "$new_firstname$($new_lastname.substring(0,$count))"
#    } until ( ! (Get-aduser -Filter {SamAccountName -eq $new_samaccountname} ) )
#
#    $new_samaccountname.ToLower()
#} 

### Otherwise FirstName.LastName
$new_samaccountname = ($new_displayname -replace " ", ".")

##### Check New User Exists #####

$NewUser = $(try {Get-ADUser $new_samaccountname -Properties * | Select-Object Name} catch {$null})
 
If ($NewUser -eq $Null) {
    "Username doesn't exist, creating user"

}
Else {
    "Username already exists, choosing next letter in sequence"
}

##### Set Username #####
$new_samaccountname = Generate-username -firstname $new_firstName -lastname $new_lastname

Write-Host "Username will be $new_samaccountname"

#####Create account by copying user or user's manager's basic attributes#####

$copy = if ($copy = $ad_account_to_copy ) {
    # return copy
    $copy
}
elseif ($copy = $ad_account_manager ) {
    # return copy
    $copy
}
else {
    Write-Host 'No manager or user specified'
    # cannot create user no copy returned
}

$path = if ($path = $CopyPath ) {
    $path
}
elseif ($path = $ManagerPath ) {
    $path
}
else {
    Write-Host 'No Path found'
}

if ($copy) { 
    $Parameters = @{
        'SamAccountName' =  $new_samaccountname 
        'Instance' = $copy 
        'Name' = $new_name 
        'DisplayName' = $new_displayname 
        'GivenName' = $new_firstname 
        'Surname' = $new_lastname 
        'PasswordNeverExpires' = $password_never_expires 
        'CannotChangePassword' = $cannot_change_password 
        'EmailAddress' = ($new_firstname + '.' + $new_lastname + '@' + "domain.com.au") 
        'Enabled' = $enable_user_after_creation 
        'UserPrincipalName' = ($new_samaccountname + '@' + "domain.com.au") 
        'AccountPassword' = (ConvertTo-SecureString -AsPlainText $Password -Force) 
        'Path' = $path
    # other changes   
            }
New-ADUser @Parameters
}
else {
    Write-Host 'No account was specified.'
}

##### Confirm new account created

Get-ADUser $new_samaccountname

## Mirror all the groups of original account that is a member of or only copy Distribution Lists

#Option 1

If ($samaccount_to_copy) {
    $UserGroups = @()
    $UserGroups = (Get-ADUser -Identity $samaccount_to_copy -Properties MemberOf).MemberOf
    Foreach ($Group in $UserGroups) {
        Add-ADGroupMember -Identity $Group -Members $new_samaccountname
        $UserGroups = (Get-ADUser -Identity $new_samaccountname -Properties MemberOf).MemberOf
    }
} 
Else { 
    ($manageraccount)
    Write-Host "Please Add Group Memberships Manually" 
#    ForEach ($Group in Get-DistributionGroup) { 
#        ForEach ($Member in Get-DistributionGroupMember -identity $Group.Name | Where-Object { $_.Name -eq $manageraccount }) { 
#            $Group.name 
#        } 
#    }
#    Add-DistributionGroupMember -Identity $Group.Name -Member $new_samaccountname -verbose
}


#Option 2 (Copies all groups filtered by ones only in specified OU (Distribution Groups)

#Get-ADPrincipalGroupMembership -Identity $manageraccount | 
#Where-Object {$_.DistinguishedName -match 'OU=Distribution Groups,OU=PRG,DC=iwf,DC=com,DC=au'}
#}

Set-ADUser -Identity $new_samaccountname -Add @{ProxyAddresses = "SMTP:" + $new_samaccountname + '@' + "domain.com.au"}
Set-ADUser -Identity $new_samaccountname -Add @{ProxyAddresses = "SMTP:" + $new_firstname + '.' + $new_lastname + '@' + "domain.com.au"}

### Set EmployeeID and Account Expirey if not set
#If ($empCode -notin ("TEMP","CONTRACTOR","NONPAYG") -AND $ad_account_to_copy -ne $NULL -or $ad_account_manager -ne $null) 
#{
#    $empCode = $vb::inputbox("Enter Employee ID")
#    Set-ADUser $new_samaccountname -EmployeeID $empCode 
#    Get-ADUser $new_samaccountname | Set-ADAccountExpiration -timespan 30.0:0
#} 
#Else 
#{
#"No Account Specified"
#}

##### Connect to Different Server if office365 hosted on that to use Office365 PowerShell
#$O365Session = New-PSSession -ComputerName
#Invoke-Command –Session $O365Session –ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta ; Import-Module MsOnline ; Connect-MsolService}
#Import-PSSession –Session $O365Session –allowclobber 

##### Get Office365 Credentials and Connect to Office365 PowerShell#####
$Office365_Credentials = Get-Credential
Import-Module MsOnline
Connect-MsolService -Credential $Office365_Credentials

##### Get License Details #####
Get-MsolAccountSku

##### Set Office365 License #####
Set-MsolUser -UserPrincipalName "$new_samaccountname@domain.com.au" -UsageLocation "AU"
Set-MsolUserLicense -UserPrincipalName "$new_samaccountname@domain.com.au"  -AddLicenses "Company:ENTERPRISEPACK"
