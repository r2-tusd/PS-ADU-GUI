Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

#Varibales
    $GFont                      = 'Consolas,10'
    $Xrow1                      = 20
    $Xrow2                      = 230
    $outpath                    = '\\10.153.2.222\c$\temp'
    $MakeAccountScriptPath      = ''

<#BEGIN
Form and elements =========================================================================
#>

#FORM
    $Form                            = New-Object system.Windows.Forms.Form
    $Form.ClientSize                 = '470,350'
    $Form.text                       = "Make New User CSV"
    $Form.TopMost                    = $True
    $Form.MinimizeBox                = $False
    $Form.Maximizebox                = $False
    $Form.FormBorderStyle            = 'FixedSingle'
#END FORM

#ELEMENTS - left to right, Top down

#Textbox legal first name
    $tbLFN                          = New-Object system.Windows.Forms.TextBox
    $tbLFN.location                 = New-Object System.Drawing.Point($Xrow1,20)
    $tbLFN.multiline                = $false
    $tbLFN.width                    = 200
    $tbLFN.height                   = 20
    $tbLFN.Font                     = $GFont
    $tbLFN.Text                     = ""

#Label legal first Name
    $labLFN                         = New-Object system.Windows.forms.Label
    $LabLFN.location                = New-Object System.Drawing.Point($Xrow2,20)
    $labLFN.width                   = 200
    $labLFN.height                  = 20
    $LabLFN.Font                    = $GFont
    $LabLFN.Text                    = "Legal First Name"

#Textbox legal last name
    $tbLLN                          = New-Object system.Windows.Forms.TextBox
    $tbLLN.location                 = New-Object System.Drawing.Point($Xrow1,50)
    $tbLLN.multiline                = $false
    $tbLLN.width                    = 200
    $tbLLN.height                   = 20
    $tbLLN.Font                     = $GFont
    $tbLLN.Text                     = ""

#label legal last name
    $labLLN                         = New-Object system.Windows.forms.Label
    $labLLN.location                = New-Object System.Drawing.Point($Xrow2,50)
    $labLLN.width                   = 200
    $labLLN.height                  = 20
    $labLLN.Font                    = $GFont
    $labLLN.Text                    = "Legal Last Name"

#Textbox Nickname
    $tbNN                          = New-Object system.Windows.Forms.TextBox
    $tbNN.location                 = New-Object System.Drawing.Point($Xrow1,80)
    $tbNN.multiline                = $false
    $tbNN.width                    = 200
    $tbNN.height                   = 20
    $tbNN.Font                     = $GFont
    $tbNN.Text                     = ""

#label Nickname
    $labNN                         = New-Object system.Windows.forms.Label
    $labNN.location                = New-Object System.Drawing.Point($Xrow2,80)
    $labNN.width                   = 200
    $labNN.height                  = 20
    $labNN.Font                    = $GFont
    $labNN.Text                    = "Nickname"

#Textbox job title
    $tbJT                          = New-Object system.Windows.Forms.TextBox
    $tbJT.location                 = New-Object System.Drawing.Point($Xrow1,120)
    $tbJT.multiline                = $false
    $tbJT.width                    = 200
    $tbJT.height                   = 20
    $tbJT.Font                     = $GFont
    $tbJT.Text                     = ""

#label job title
    $labJT                         = New-Object system.Windows.forms.Label
    $labJT.location                = New-Object System.Drawing.Point($Xrow2,120)
    $labJT.width                   = 200
    $labJT.height                  = 20
    $labJT.Font                    = $GFont
    $labJT.Text                    = "Job Title"

#Textbox site
    $tbSi                          = New-Object system.Windows.Forms.TextBox
    $tbSi.location                 = New-Object System.Drawing.Point($Xrow1,160)
    $tbSi.multiline                = $false
    $tbSi.width                    = 200
    $tbSi.height                   = 20
    $tbSi.Font                     = $GFont
    $tbSi.Text                     = ""

#label site
    $labSi                         = New-Object system.Windows.forms.Label
    $labSi.location                = New-Object System.Drawing.Point($Xrow2,160)
    $labSi.width                   = 200
    $labSi.height                  = 20
    $labSi.Font                    = $GFont
    $labSi.Text                    = "Site"

#Textbox Department
    $tbDe                          = New-Object system.Windows.Forms.TextBox
    $tbDe.location                 = New-Object System.Drawing.Point($Xrow1,190)
    $tbDe.multiline                = $false
    $tbDe.width                    = 200
    $tbDe.height                   = 20
    $tbDe.Font                     = $GFont
    $tbDe.Text                     = ""

#label Department
    $labDe                         = New-Object system.Windows.forms.Label
    $labDe.location                = New-Object System.Drawing.Point($Xrow2,190)
    $labDe.width                   = 200
    $labDe.height                  = 20
    $labDe.Font                    = $GFont
    $labDe.Text                    = "Department"

#Textbox Employee Number
    $tbEN                          = New-Object system.Windows.Forms.TextBox
    $tbEN.location                 = New-Object System.Drawing.Point($Xrow1,230)
    $tbEN.multiline                = $false
    $tbEN.width                    = 200
    $tbEN.height                   = 20
    $tbEN.Font                     = $GFont
    $tbEN.Text                     = ""

#label Employee Number
    $labEN                         = New-Object system.Windows.forms.Label
    $labEN.location                = New-Object System.Drawing.Point($Xrow2,230)
    $labEN.width                   = 200
    $labEN.height                  = 20
    $labEN.Font                    = $GFont
    $labEN.Text                    = "Employee #"

#Textbox Manager Email
    $tbME                          = New-Object system.Windows.Forms.TextBox
    $tbME.location                 = New-Object System.Drawing.Point($Xrow1,270)
    $tbME.multiline                = $false
    $tbME.width                    = 200
    $tbME.height                   = 20
    $tbME.Font                     = $GFont
    $tbME.Text                     = ""

#label Manger Email
    $labME                          = New-Object system.Windows.forms.Label
    $labME.location                 = New-Object System.Drawing.Point($Xrow2,270)
    $labME.width                    = 200
    $labME.height                   = 20
    $labME.Font                     = $GFont
    $labME.Text                     = "Manager Email"


#label file saved
    $labFS                          = New-Object system.Windows.forms.Label
    $labFS.location                 = New-Object System.Drawing.Point($Xrow1,310)
    $labFS.width                    = 200
    $labFS.height                   = 20
    $labFS.Font                     = $GFont
    $labFS.Text                     = "File Saved!"
    $labFS.Visible                  = $False



#button Make CSV
    $btnMCSV                        = New-Object system.Windows.Forms.Button
    $btnMCSV.location               = New-Object System.Drawing.Point($Xrow1,310)
    $btnMCSV.width                  = 200
    $btnMCSV.height                 = 24
    $btnMCSV.Font                   = $GFont
    $btnMCSV.text                   = "Make CSV"
    $btnMCSV.Enabled                = $true
    $btnMCSV.Visible                = $true

#button Reset
    $btnRe                          = New-Object system.Windows.Forms.Button
    $btnRe.location                 = New-Object System.Drawing.Point($Xrow2,310)
    $btnRe.width                    = 90
    $btnRe.height                   = 24
    $btnRe.Font                     = $GFont
    $btnRe.text                     = "Reset"

#button Make Accounts
    $btnMA                          = New-Object system.Windows.Forms.Button
    $btnMA.location                 = New-Object System.Drawing.Point(340,310)
    $btnMA.width                    = 90
    $btnMA.height                   = 24
    $btnMA.Font                     = $GFont
    $btnMA.text                     = "Make Acct"

<#END
Form and elements =========================================================================
#>

#Controls - If its not listed here you can't interact with it on the form
#region Controls
$Form.controls.AddRange(@(
    $labLFN,$tbLFN,`
    $labLLN,$tbLLN,`
    $labNN,$tbNN,`
    $labJT,$tbJT,`
    $labSi,$tbSi,`
    $labDe,$tbDe,`
    $labEN,$tbEN,`
    $labME,$tbME,`
    $btnMCSV,$btnRe,$labFS,$btnMA
    ))
#endregion Controls

#Events
#region Events
$btnMCSV.Add_MouseClick({Make-CSV})
$btnRe.Add_MouseClick({Reset-Form})
$btnMA.Add_MouseClick({Make-accounts})
#$tbLFN.Add_keyDown({if ($_.KeyCode -eq "Enter") {}})
#endregion Events

#Functions
#region Functions
Function Make-CSV {
    Hide-Button
    $date = get-date -format yyyyMMddHHmmss
    $outfile = "NewUsers-$date.csv"
    
    $NewCSV = {} | Select EmployeeNum,LegalFirstName,Nickname,LegalLastName,JobTitle,Site,Department,CertClass,Man,ManEmail | Export-Csv $outpath\$outfile
    
    $User = Import-Csv $outpath\$outfile
        $User.EmployeeNum       = $tbEN.text
        $User.LegalFirstName    = $tbLFN.text
        $User.LegalLastName     = $tbLLN.text
        $User.Nickname          = $tbNN.text
        $User.JobTitle          = $tbJT.text
        $User.Site              = $tbSi.text
        $User.Department        = $tbDe.text
        $user.ManEmail          = $tbME.text
    $User | Export-Csv $outpath\$outfile
   
    ii $outpath
    Check-File

    
}

Function Hide-Button {
    $btnMCSV.Enabled        = $false
    $btnMCSV.Visible        = $false
}

Function Reset-Form {
    $tbEN.clear()
    $tbLFN.Clear()
    $tbLLN.Clear()
    $tbNN.Clear()
    $tbJT.Clear()
    $tbSi.Clear()
    $tbDe.Clear()
    $tbME.Clear()
    $btnMCSV.enabled        = $true
    $btnMCSV.Visible        = $true
    $labFS.Visible          = $False
}

Function Check-File {
    Start-Sleep 3
    IF(test-path $outpath\$outfile) {$labFS.Visible = $true}
}

Function Make-accounts {    
    start-process pwsh.exe -ArgumentList "-interactive -file C:\scripts\MasterUserCreatorV2.ps1" -UseNewEnvironment -wait -WindowStyle Maximized
}

#endregion Functions

#endregion Master

[void]$Form.ShowDialog()
