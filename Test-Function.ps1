Function Variables {

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

$Variables = @($samaccount_to_copy,`
               $manageraccount,`
               $new_displayname,`
               $new_firstname,`
               $new_lastname,`
               $new_name,`
               $CopyPath,`
               $ManagerPath,`
               $enable_user_after_creation,`
               $password_never_expires,`
               $cannot_change_password,`
               $ad_account_to_copy,`
               $ad_account_manager)
               }

Function DinoPassPassword
{
$web = New-Object Net.WebClient #Generates powershell web client
$web.Headers.Add("Cache-Control", "no-cache");
$PwdString = $web.DownloadString("http://www.dinopass.com/password/simple")
$PwdString = $PwdString.substring(0, 1).toUpper() + $PwdString.substring(1)
$Password = ConvertTo-SecureString -String $PwdString -AsPlainText -Force 
}

### If not using DinoPass ###
Function random-password ($length = 8) {
    $punc = 46..46
    $digits = 48..57
    $letters = 65..90 + 97..122

    # Thanks to
    # https://blogs.technet.com/b/heyscriptingguy/archive/2012/01/07/use-pow
    $password = get-random -count $length `
        -input ($punc + $digits + $letters) |
        % -begin { $aa = $null } `
        -process {$aa += [char]$_} `
        -end {$aa}

    return $password
}

##### Generate Username -FLastName
Function FLastName($new_displayname)
{
    $count = 0

    do
   {
        $count++
        $new_samaccountname = "$($new_firstname.substring(0,$count))$new_lastname"
    } until ( ! (Get-aduser -Filter {SamAccountName -eq $new_samaccountname} ) )

    $new_samaccountname.ToLower()
}  

##### Generate Username - FirstNameFirstInitialLastName #####
Function FirstNameL($new_displayname)
{
    $count = 0

    do
    {
        $count++
        $new_samaccountname = "$new_firstname$($new_lastname.substring(0,$count))"
    } until ( ! (Get-aduser -Filter {SamAccountName -eq $new_samaccountname} ) )

    $new_samaccountname.ToLower()
} 

Function FirstName.LastName
{
$new_samaccountname = ($new_displayname -replace " ", ".")
}

Function Domain{
Param([string]$Domain = 'Select Domain'
)
cls
Write-Host "==========$Domain=========="

Write-Host "1: Press 1 for Domain1."
Write-Host "2: Press 2 for Domain2."
Write-Host "3: Press 3 for Domain3."
Write-Host "Q: Press Q to quit."
}

do
{
    Variables
    Domain
    $input = Read-Host "Please select domain"
    switch ($input)
    {
        '1'{
            cls
            'Domain1'
            $Password = DinoPassPassword
            $new_samaccountname = FLastName            
        }'2'{
            cls
            'Domain2'
            $Password = DinoPassPassword
            $new_samaccountname = FirstName.LastName
            
        }'3'{
            cls
            'Domain3'
            $Password = random-password
            $new_samaccountname = FirstNameL
            
        }'q'{
           return
        }
     }
     pause
  }
until ($input -eq 'q')