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
# Need    : 7-Zip, Powershell, grindFile.ps1, 
#           grindFolder.ps1
###########################################################
# PARAMETERS
###########################################################

$dateStamp = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)

#The full path of the log file
$logFile = "D:\Excel Link Grinder\log.txt"

#Array of file names
$Files = New-Object System.Collections.ArrayList

#Array of ints that represent he number of changes made to a document
$NumChanges = New-Object System.Collections.ArrayList

#Array of files that have been changed
$LinksChanged = New-Object System.Collections.ArrayList

###########################################################
# FUNCTIONS
###########################################################
. "D:\Excel Link Grinder\Main\grindFile.ps1"
. "D:\Excel Link Grinder\Main\grindFolder.ps1"
. "D:\Excel Link Grinder\Main\Grinder_GUI.ps1"
###########################################################
# MAIN FUNCTION
###########################################################

if($Files.Count -eq 0)
{
    Write-Warning "An error occured trying to find this file." 
}

###########################################################
# LOG
###########################################################

foreach($item in $Files)
{
    if($userRep -eq 1)
    {
        $addString = $dateStamp + "  Original file:" + $Files[0] + "`t`t" + "New file:" + $Files[1] + " `t`t" + "External links changed: "
        foreach($j in $LinksChanged)
        {
            $addString = $addString + $j + " "
        }
        Add-Content $logFile -Value $addString -Force 
    }
    else
    {
        try
        {
            $addString = $dateStamp + "  File scanned: " + $item +  "`t`t" + "# of changes made: " + $NumChanges.Item([array]::IndexOf($Files,$item)) 
        }
        catch
        {
            $addString = $dateStamp + "  File scanned: " + $item +  "`t`t" + "# of changes made: 0 "
        }
       
        
        Add-Content $logFile -Value $addString -Force
    }
          
}

###########################################################
#
###########################################################