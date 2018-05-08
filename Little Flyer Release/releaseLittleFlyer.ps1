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
	Write-Output "Date variables have been set."

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
	Write-Output "Working year value set as $workflowYear"
	Write-Output "Working Month value set as $workflowMonth"	

# Setting our script home path variable
	try {
		# Script Home Directory
			$scriptPath = (Get-Item -Path ".\").FullName
	}
	catch {
		myErrorMessage("Error setting home path variables")
	}
	Write-Output "Script path set."

# Setting our PDF home path variables
	try {
		# PDF Directories
			$pdfHomeDirectory = "$scriptPath\PDFs"
			$pdfWorkDirectory = "$pdfHomeDirectory\workDir"
	}
	catch {
		myErrorMessage("Error setting PDF home path variables")
	}
	Write-Output "PDF directories set."
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
	Write-Output "Workflow month path set."

# Begin process of photoshop work
	Write-Output "Starting pdf copy process for month."

	# Ask for PSD filename to copy to PDF work directory
	Function Get-FileName($initialDirectory)
		{
			[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
			$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
			$OpenFileDialog.initialDirectory = $initialDirectory
			$OpenFileDialog.filter = "Pdf Files|*.pdf|All files (*.*)|*.*"
			$OpenFileDialog.ShowDialog() | Out-Null
			$OpenFileDialog.filename
		}
	try {
        Write-Output "To start, select monthly PDF copy."
        pause
		$initialPdfFile = Get-FileName $scriptPath
		}
	catch {
		myErrorMessage("Error getting the monthly pdf file to start copying")
	}

	# Copy the monthly pdf file to work file

		Write-Output "Copying pdf file to work directory."
		$pdfWorkFile = '$$$$_LF_Hi.pdf'
		try {
			Copy-Item -Path $initialPdfFile -Destination "$pdfWorkDirectory\$pdfWorkFile" -Force
		}
		catch {
			myErrorMessage("Error copying the monthly pdf file to work directory")
		}

	# Alert user to open photoshop and export files
	Write-Output "Open Photoshop and export Little Flyer images to work directory.  Press any key to continue once work files have been successfully exported."
	pause

	
	# Perform photoshop tasks

	$photoshopWork = "Begin Photoshop image resize process for full size files",
		"Begin Photoshop image resize process for min files"
	$onMin = $false
	Foreach ($step in $photoshopWork)
	{
		Write-Output $step
		pause
		Write-Output "Copying files"

		# Copy the files exported to appropriate directory
			try {
				Get-ChildItem -Path $pdfWorkDirectory -Filter "*.jpg" | ForEach-Object {
				Copy-Item $_.FullName -Destination $workflowMonthDirectory
				}
			}
			catch {
				myErrorMessage("Error copying image to month workflow directory")
			}

		# Rename the files
			try {
				if ($onMin -eq $true) {
					$files = Get-ChildItem -Path $workflowMonthDirectory -Exclude "*$workflowYear*"
				} else {
					$files = Get-ChildItem -Path $workflowMonthDirectory -Exclude "*min*"
				}
			}
			catch {
				myErrorMessage("Error getting image items in workflow month directory")
			}
			foreach ($file in $files) 
			{
				try {
					$curFileName = $file.BaseName
					$curFileExt = $files.Extension
					if ($onMin -eq $false) {
						$newFileName=$file.Name.Replace($curFileName,"LF-$workflowMonth-$workflowYear-$curFileName")
					} else {
						$newFileName=$file.Name.Replace($curFileName,"LF-$workflowMonth-$workflowYear-$curFileName-min")
					}
					Rename-Item $file $newFileName
					$curFileName = ""
					$curFileExt = ""
				}
				catch {
					myErrorMessage("Error renaming files")
				}
			}
		if ($onMin -eq $false) {
			$onMin = $true
		}
	}

	# Work complete.  If no errors, ask if the user wants the file opened and to visit tinypng where we compress the image files.

	if (!$Error) {
		$openSaveLocation = $popup.Popup("Work completed without any errors detected.'n'nWould you like to open the directory in explorer?",0,"Completed Successfully!",4)
		if ($openSaveLocation -eq 6) {
			start https://tinypng.com
			ii $workflowMonthDirectory
		} else {
			exit
		}
	}


# Catch errors and report their message
function myErrorMessage {
    param( [string]$myErrMsg )
    if ( $popup.Popup("Houston, we have an error: $myErrMsg.  Should I stop?",0,"Error Detected",4) -eq 6 ) {
        exit
    } else {
		$error.clear()
    }
}