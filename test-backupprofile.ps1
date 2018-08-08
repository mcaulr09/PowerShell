$destination = "\\qnap-451\MCAULR09\"

$folder = "Desktop",
"Downloads",
"Favorites",
"Documents",
"Music",
"Pictures",
"Videos",
"AppData\Local\Mozilla",
"AppData\Local\Google"

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
		$currentLocalFolder = $userprofile + "\" + $f
		$currentRemoteFolder = $backupfolder + $f
		robocopy $currentLocalFolder $currentRemoteFolder /E /XO /NP /LOG+:"$backupfolder\log.txt"
	}
}
	
