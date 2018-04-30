#
# Developer: Ryan Valizan
# Project: Little Flyer Automated Release
# Purpose: Rename all the Little Flyer images and put them in their folder to upload to the server.
#

# clear our error count
	if ($error.Count -gt 0) {
		$error.clear()
	}

# Setting up our date variables

	# Set month variables
	try {
		$curMonthAbv = ((Get-Date).AddMonths(0)).ToString('MMM')
		$nextMonthAbv = ((Get-Date).AddMonths(1)).ToString('MMM')
		$nextMonth = ((Get-Date).AddMonths(1)).ToString('MMMM')
	}
	catch {
		myErrorMessage("Check date variables")
	}

	# determin which month to work with based on workflow needs and set the month variables
	try {
		$popup = new-object -comobject wscript.shell 
		$popupMonth = $popup.popup("Is this for the $nextMonth Little Flyer?",0,"When are you Publishing this?",4) 
	}
	catch {
		myErrorMessage("Error setting the workflow month")
	}

	# Determin the year and month our workflow requires
    try {
		If ($popupMonth -eq 6) { 
			$workflowMonth = $nextMonthAbv
			If (((Get-Date).Month) -eq 12) {
				$workflowYear = ((Get-Date).AddYears(1)).ToString('yyyy')
			} else {
				$workflowYear = ((Get-Date).AddYears(0)).ToString('yyyy')
			}
		} else { 
			$workflowMonth = $curMonthAbv
			$workflowYear = ((Get-Date).AddYears(0)).ToString('yyyy')
		}
    }
    catch {
        myErrorMessage("Error setting the workflow month and year required")
    }

# Setting our script home path variable
	try {
		# Script Home Directory
			$scriptPath = (Get-Item -Path ".\").FullName
	}
	catch {
		myErrorMessage("Error setting home path variables")
	}
# Setting our PDF home path variables
	try {
		# PDF Directories
			$pdfHomeDirectory = "$scriptPath\PDFs"
			$pdfWorkDirectory = "$pdfHomeDirectory\workDir"
	}
	catch {
		myErrorMessage("Error setting PDF home path variables")
	}
# Setting our output home path variables
	# Set output path year variable and create directory
	try {
		# Output Directories
			$workflowYearDirectory = "$scriptPath\$workflowYear"

		# Test if output year directory exists.  If not, create it
			If(!(test-path $workflowYearDirectory))
				{
					New-Item -ItemType directory -Force -Path $workflowYearDirectory
				}
	}
	catch {
		myErrorMessage("Error setting workflow year home path variable")
	}
	# Set output path month variable and create directory
	try {
		# Output Directories
			$workflowMonthDirectory = "$workflowYearDirectory\$workflowMonth"

		# Test if output month directory exists.  If not, create it
			If(!(test-path $workflowMonthDirectory))
				{
					New-Item -ItemType directory -Force -Path $workflowMonthDirectory
				}
	}
	catch {
		myErrorMessage("Error setting workflow month home path variable")
	}

