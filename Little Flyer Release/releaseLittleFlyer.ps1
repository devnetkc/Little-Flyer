#
# Developer: Ryan Valizan
# Project: Little Flyer Automated Release
# Purpose: Rename all the Little Flyer images and put them in their folder to upload to the server.
#

# clear our error count
if ($error.Count -gt 0) {
    $error.clear()
}

# try setting up our date variables
try {
    $a = new-object -comobject wscript.shell 
    $curMonthAbv = ((Get-Date).AddMonths(0)).ToString('MMM')
    $nextMonthAbv = ((Get-Date).AddMonths(1)).ToString('MMM')
    $nextMonth = ((Get-Date).AddMonths(1)).ToString('MMMM')

	$gDate = Get-Date
	$cYR = $gDate.ToString('yyyy')
	$cMON = $gDate.ToString('MMM')
	$saveLoc = 'C:\Users\ryan\OneDrive - cfm Distributors, Inc\_Website\Wordpress\cfm_live\images\Little Flyer\' + $cYR
	$workDir = $saveLoc+'\workDir'
	$newDir = $saveLoc + '\' + $cMON
	$copyFrom = $workDir + '\*' 
}
catch {
    myErrorMessage(17)
}

If(!(test-path $newDir))
    {
        New-Item -ItemType directory -Force -Path $newDir
    }

Copy-Item $copyFrom $newDir
cd $newDir
get-childitem | ForEach-Object { Move-Item -LiteralPath $_.name $_.name.Replace('$$$',"$cMON")}
