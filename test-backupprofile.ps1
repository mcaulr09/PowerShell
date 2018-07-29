$destination = "\\qnap-451\MCAULR09\"

$folder = "Desktop",
"Downloads",
"Favorites",
"Documents",
"Music",
"Pictures",
"Videos",
"AppData\Local\Mozilla",
"AppData\Local\Google",
"AppData\Roaming\Mozilla"

###############################################################################################################

$username = $env:username
$userprofile = $env:userprofile
$appData = $env:localAPPDATA

$backupfolder = Join-Path $destination "Profile\"

$TestPath = Test-Path $backupfolder
if ($TestPath -eq $false)
{
    New-Item $backupfolder
    }
Else
{

###### Backup Data section ########
	
	foreach ($f in $folder)
	{	
		$currentLocalFolder = $userprofile + $f
		$currentRemoteFolder = $backupfolder + $f

$GCI_Params = @{
    ErrorAction = 'silentlyContinue'
    Path = $currentLocalFolder
    Recurse = $True
    Force = $True
    }
$MO_Params = @{
    ErrorAction = 'silentlyContinue'
    Property = 'Length'
    Sum = $True
    }
$currentFolderSize = (Get-ChildItem @GCI_Params |
    Measure-Object @MO_Params ).
    Sum / 1MB
		robocopy $currentLocalFolder $currentRemoteFolder /E /XO /NP
	}
}
	
