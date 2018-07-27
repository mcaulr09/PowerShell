$destination = "\\qnap-451\MCAULR09"

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

$username = gc env:username
$userprofile = gc env:userprofile
$appData = gc env:localAPPDATA

([IO.Directory]::Exists($destination + "\" + "Profile" + "\"))

###### Backup Data section ########
	write-host -ForegroundColor green "Backing up data from local machine for $username"
	
	foreach ($f in $folder)
	{	
		$currentLocalFolder = $userprofile + "\" + $f
		$currentRemoteFolder = $destination + "\" + "Profile" + "\" + $f
		$currentFolderSize = (Get-ChildItem -ErrorAction silentlyContinue $currentLocalFolder -Recurse -Force | Measure-Object -ErrorAction silentlyContinue -Property Length -Sum ).Sum / 1MB
		$currentFolderSizeRounded = [System.Math]::Round($currentFolderSize)
		write-host -ForegroundColor cyan "  $f... ($currentFolderSizeRounded MB)"
		robocopy $currentLocalFolder $currentRemoteFolder /S /E /XO
	}
	
