<# 
.NAME
    Chase Dudas
#>

. "D:\Excel Link Grinder\Main\grindFile.ps1"
. "D:\Excel Link Grinder\Main\grindFolder.ps1"

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#region begin GUI{ 

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '400,400'
$Form.text                       = "Excel Link Grinder"
$Form.TopMost                    = $false
$Form.StartPosition              = "CenterScreen"
$Form.AutoScale                  = $true
$Form.MaximizeBox                = $false
$Form.MinimizeBox                = $false
$Form.BackColor                  = "#c9e2c3"

$Button1                         = New-Object system.Windows.Forms.Button
$Button1.text                    = "File"
$Button1.width                   = 150
$Button1.height                  = 150
$Button1.location                = New-Object System.Drawing.Point(25,200)
$Button1.Font                    = 'Comic Sans MS,20,style=Bold'
$Button1.FlatStyle               = "System"

$Button2                         = New-Object system.Windows.Forms.Button
$Button2.text                    = "Folder"
$Button2.width                   = 150
$Button2.height                  = 150
$Button2.location                = New-Object System.Drawing.Point(225,200)
$Button2.Font                    = 'Comic Sans MS,20,style=Bold'
$Button2.FlatStyle               = "System"

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Click which option you wish to grind:"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 0
$Label1.location                 = New-Object System.Drawing.Point(15,75)
$Label1.Font                     = 'Comic Sans MS,15,style=Bold'

$Form.controls.AddRange(@($Button1,$Button2,$Label1))

function OnClick1
{ 
    $Form.Close()
	grindFile 
    
}
function OnClick2
{ 
    $Form.Close()
	grindFolder
    
}
#region gui events {
$Button1.Add_Click({ OnClick1 })
$Button2.Add_Click({ OnClick2 })
#endregion events }

#endregion GUI }


#Write your logic code here

[void]$Form.ShowDialog()