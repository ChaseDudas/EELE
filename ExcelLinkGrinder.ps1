###########################################################
# AUTHOR  : Chase Dudas
# CREATED : 7/3/2018
# Title   : Excel Link Grinder
# COMMENT : This script recurses through a folder given by
#           the user. It checks for external links that are
#           deemed bad by the if else statements.It also 
#           toggles on/off Read Only to make changes. 
#           This file can be manipulated to traverse entire
#           folders by removing the first if statement
#           after the for each statement. 
# Need    : 7-Zip, Powershell 
###########################################################
# PARAMETERS
###########################################################

$dateStamp = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)

#Dialog box to prompt the user for what zip file they would like to check
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$zipfileFolder = [Microsoft.VisualBasic.Interaction]::InputBox("Put the file path of the folder to grind:", "Excel Link Grinder", "D:\Excel Link Grinder\Grinder")

#Dialog box to recieve the users input for what file to search for
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$wantedFile = [Microsoft.VisualBasic.Interaction]::InputBox("Put the name of the file to grind: `n Do not include .xlsx", "Excel Link Grinder", "ex.Test_File")
$wantedFile = $wantedFile +'.xlsx'

#Dialog box to get the name chosen for the new file
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$destFile = [Microsoft.VisualBasic.Interaction]::InputBox("Put the name you would like to chnage it to: `n Do not include .xlsx", "Excel Link Grinder", "ex. New_Test_File")
$destFile = $destFile + '.xlsx'

#The full path of the log fiel
$logFile = "D:\Excel Link Grinder\Log.txt"

#Array of file names
$Files = New-Object System.Collections.ArrayList

#Array of ints that represent he number of changes made to a document
$NumChanges = New-Object System.Collections.ArrayList

#Array of files that have been changed
$LinksChanged = New-Object System.Collections.ArrayList

###########################################################
# MAIN FUNCTION
###########################################################

#Loop through the folder and look at each file inside
Foreach($zippieFile in (Get-ChildItem $zipfileFolder -Include *.xlsx -Recurse))
{
    if($zippieFile.name -eq $wantedFile)
    {
        #Sets the objects full path as its file name
        $zipfileName = $zippieFile.FullName

        #Turns Read Only to False
        Set-ItemProperty $zipfileName -name IsReadOnly -value $false

        #Grabs the ~~~.xlsx part of the file path
        $zipLeaf = Split-Path -Path $zipfileName -Leaf

        #Comment to help distinguish what files are being read 
        Write-Host "Looking at:" $zipLeaf "`n"

        #Adds the file name to an array of file names
        $Files.Add("$zipLeaf")

        #Takes the file path to the xlsx file and renames it to a zip for editing
        $tempFile = $zipLeaf -replace "xlsx", "zip"
        Rename-Item -Path $zipfileName -NewName $tempFile
        $zipfileName = $zipfileName -replace "xlsx", "zip" 

        #The prompted file name + the file path to the external links I want to edit
        $zZipfileName = $zipfileName + "\xl\externalLinks\_rels"
    
        #Counts the number of bad links in a file
        $countBadLinks = 0

        #Grabs the files that are contained in the _rels sub folder
        $shap = new-object -com shell.application 
        $zipFile = $shap.Namespace($zZipfileName) 
        $i = $zipFile.Items() | select @{n='Name'; e={Split-Path -Path $_.path -Leaf}}, @{n='Path'; e={$_.path}}

        # Open zip and find the particular file 
        Add-Type -assembly  System.IO.Compression.FileSystem
        $zip =  [System.IO.Compression.ZipFile]::Open($zipfileName,"Update")

        #For display
        Write-Host "#################################################################`n"

        #Loop to Read/Write the external links 
        Foreach ($zippie in $i)
        {
            #Read the file 
            Write-Host "Reading" $zippie.name "`n"
            $desiredFile = [System.IO.StreamReader]($zip.Entries | Where-Object {$_.FullName -match $zippie.Name}).Open()
            $text = $desiredFile.ReadToEnd()
            $desiredFile.Close()
            $desiredFile.Dispose()

            #A bool that is false if the link is good, and true if the link is bad
            $link_quality = $false

            Write-Host "Analyzing text..."
            # Looks at the $text variable to see if the link is good or bad. Sets the boolean to true if bad
            if($text -like '*\\nascfs02\gsa\ACCTNG*')
            {
                Write-Host "              ...found a match for \\nascfs02\gsa\ACCTNG. Link has been replaced.`n" -BackgroundColor Red

                #Replaces the text with the correct drive
                $text = $text.Replace("\\nascfs02\gsa\ACCTNG", "I:\Accounting")

                #This statement accounts for the camel casing of the like method
                $text = $text.Replace("\\nascfs02\GSA\ACCTNG", "I:\Accounting")

                #Sets the bool to true so the file can be written to
                $link_quality = $true

            }
            elseif($text -like '*\\nascfs02\GSA\common\ACCTNG*')
            {
                Write-Host "              ...found a match for \\nascfs02\GSA\common\ACCTNG. Link has been replaced.`n" -BackgroundColor Red

                #Replaces the text with the correct drive
                $text = $text.Replace("\\nascfs02\GSA\common\ACCTNG", "I:\Accounting")

                #Sets the bool to true so the file can be written to
                $link_quality = $true
            }
            elseif($text -like '*G:\ACCTNG*')
            {
                Write-Host "              ...found a match for G:\ACCTNG. Link has been replaced.`n" -BackgroundColor Red

                #Replaces the text with the correct drive
                $text = $text.Replace("G:\ACCTNG", "I:\Accounting")

                #Sets the bool to true so the file can be written to
                $link_quality = $true
            }
            elseif($text -like '*I:\common\ACCTNG*')
            {
                Write-Host "              ...found a match for I:\common\ACCTNG. Link has been replaced.`n" -BackgroundColor Red

                #Replaces the text with the correct drive
                $text = $text.Replace("I:\common\ACCTNG", "I:\Accounting")

                #Sets the bool to true so the file can be written to
                $link_quality = $true
            }
            else
            {
                Write-Host "              ...no matches. Skipping to next file.`n" -BackgroundColor DarkGreen
            }

            #File is only written to if the link is bad
            if($link_quality)
            {
                Write-Host "Opening the file for writing..."

                $LinksChanged.Add($zippie.name)

                #Increments the number of bad links
                $countBadLinks++

                # Open the file again and Write
                $desiredFile = [System.IO.StreamWriter]($zip.Entries | Where-Object {$_.FullName -match $zippie.Name}).Open()

                # If needed, zero out the file -- in case the new file is shorter than the old one
                $desiredFile.BaseStream.SetLength(0)

                # Insert the $text to the file and close
                $desiredFile.Write($text -join "`r`n")

                #Flush and close for the next iteration
                $desiredFile.Flush()
                $desiredFile.Close()

                Write-Host "                            ...file writing done"
            }

            #formatting 
            Write-Host "#################################################################`n"
        }

        # Closes the zip file
        $zip.Dispose()

        #Sets Read Only back to true
        Set-ItemProperty $zipfileName -name IsReadOnly -value $true

        #Renames the file back to an xlsx document
        Rename-Item -Path $zipfileName -NewName $destFile

        #Add to the NumChanges file
        $NumChanges.Add("$countBadLinks") 

        Write-Host "~~~ Finished ~~~`n"
    }
    
}

if($Files.Count -eq 0)
  {
      Write-Host "Error: File not found" -BackgroundColor Yellow -ForegroundColor Black
  }

###########################################################
# END OF FUNCTION
###########################################################

###########################################################
# OUTPUT
###########################################################
#for($i = 0; $i -lt $Files.Count; $i++)
#{
#    if($i%2 -eq 0)
#    {
#        Write-Host "Name:"$Files.Item($i) "`t # of bad links: "$NumChanges.Item($i)    -BackgroundColor White -ForegroundColor Black 
#    }
#    else
#    {
#        Write-Host "Name:"$Files.Item($i) "`t # of bad links: "$NumChanges.Item($i)    -BackgroundColor Green -ForegroundColor Black     
#    }     
#}
###########################################################
# END OUTPUT
###########################################################

###########################################################
# LOG
###########################################################

foreach($item in $Files)
{
    $addString = $dateStamp + "  Original file:" + $wantedFile + "`t`t" + "New file:" + $destFile + " `t`t" + "External links changed: "
    foreach($j in $LinksChanged)
    {
        $addString = $addString + $j + " "
    }
    Add-Content $logFile -Value $addString -Force 
}

###########################################################
# END LOG
###########################################################