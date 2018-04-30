#
# Developer: Ryan Valizan
# Project: Little Flyer Automated Release
#

$gDate = Get-Date
$cYR = $gDate.ToString('yyyy')
$cMON = $gDate.ToString('MMM')
$saveLoc = 'C:\Users\ryan\OneDrive - cfm Distributors, Inc\_Website\Wordpress\cfm_live\images\Little Flyer\' + $cYR
$workDir = $saveLoc+'\workDir'
$newDir = $saveLoc + '\' + $cMON
$copyFrom = $workDir + '\*' 
If(!(test-path $newDir))
    {
        New-Item -ItemType directory -Force -Path $newDir
    }

Copy-Item $copyFrom $newDir
cd $newDir
get-childitem | ForEach-Object { Move-Item -LiteralPath $_.name $_.name.Replace('$$$',"$cMON")}
