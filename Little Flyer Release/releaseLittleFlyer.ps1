#
# Developer: Ryan Valizan
# Project: Little Flyer Automated Release
#

$curYr = Get-Date -format YYYY
$curMon = Get-Date -format MMM
$saveLoc = "C:\Users\ryan\OneDrive - cfm Distributors, Inc\_Website\Wordpress\cfm_live\images\Little Flyer\" + $curYr
$workDir = "$($saveLoc)\workDir"
$newDir = $saveLoc+$curMon

Copy-Item "$($saveLoc)\*" $newDir
dir $saveLoc | rename-item -NewName {$_.name -replace "$$$",$curMon}