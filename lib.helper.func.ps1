# ==============================================================================================
# THIS SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
# OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
# FITNESS FOR A PARTICULAR PURPOSE.
# ==============================================================================================



$CsrFilter = "Request Files (*.csr,*.req)|*.csr;*.req|All Files (*.*)|*.*"

## loading .net classes needed
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[Windows.Forms.Application]::EnableVisualStyles()
#
# define fonts and colors
$SuccessFontColor = "Green"
$WarningFontColor = "Yellow"
$FailureFontColor = "Red"

$SuccessBackColor = "Black"
$WarningBackColor = "Black"
$FailureBackColor = "Black"

$FontStdt = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Regular)
$FontBold = New-Object System.Drawing.Font("Arial",11,[System.Drawing.FontStyle]::Bold)
$FontItalic = New-Object System.Drawing.Font("Arial",9,[System.Drawing.FontStyle]::Italic)
$Icon = [system.drawing.icon]::ExtractAssociatedIcon("C:\Windows\System32\WindowsPowershell\v1.0\powershell.exe")


function Get-AdEnrollCaList
{
    $AdCaInfo = certutil -adca
    $CaData = ""|select CaName, DnsHostName 
    $CaList = @()

    foreach($line in $AdCaInfo) 
    {
        #$line
        if (($line -match "CAIsValid: 1") -or ($line -match "CertUtil: -ADCA command completed successfully")) {
            if ($CaData.CaName -ne $null) {
                $CaList += $CaData
                $CaData = ""|select CaName, DnsHostName 
            }
        }
        if ($line -match "cn =") {
           $CaData.CaName = $line.split("=")[1].trim()
        }
        if ($line -match "dNSHostName =") {
           $CaData.DnsHostName = $line.split("=")[1].trim()
        }
    }

    return $CaList
}


function Get-CATemplates
{
    param(
        [Parameter(mandatory=$True)][String]$TargetCACnfg 
    )

    $aTemplates = @()
    $templates = (certutil -config $TargetCACnfg -catemplates)#.split(":")[0]
    foreach($tmpl in $templates){
        if( $tmpl -notmatch "CertUtil"){
            $aTemplates += $tmpl.split(":")[0]
        }
    }
    return $aTemplates
}


function IsCAResponding
{
    param(
        [Parameter(mandatory=$True)][String]$TargetCACnfg 
    )

    $ret = $True
    $result = ""
    $result = certutil -config $TargetCACnfg -ping
    if( $result -match "FAILED" ) {
        $ret = $False
    }
    return $ret
}

function HasCAAdmin
{
    param(
        [Parameter(mandatory=$True)][String]$TargetCACnfg 
    )

    $ret = $True
    $result = ""
    $result = certutil -config $TargetCACnfg -pingadmin
    if( $result -match "FAILED" ) {
        $ret = $False
    }
    return $ret
}

function Select-File
{
    param(
        [Parameter(mandatory=$False)]
            $StartFolder = (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent),
        [Parameter(mandatory=$False)]
            $FileFilter = "All Files (*.*)|*.*"
    )

    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = $StartFolder 
        Filter = $FileFilter
    }
    [void] $FileBrowser.ShowDialog()
    return ($FileBrowser.filename)
}

function Select-Folder
{
    param(
        [Parameter(mandatory=$False)]
            $StartFolder = (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)
    )

    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{ 
        SelectedPath  = $StartFolder 
    }
    [void] $FolderBrowser.ShowDialog()
    return ($FolderBrowser.SelectedPath)
}

function Break-MessageBox
{
    param(
        [Parameter(mandatory=$true)]$Message
    )
   [void][System.Windows.Forms.MessageBox]::Show($Message,"Critical Error!","OK",[System.Windows.Forms.MessageBoxIcon]::Stop)
    exit
}

function DropDownBox
{
    param(
        [Parameter(mandatory=$true)][Array]$DropDownArray,
        [Parameter(mandatory=$true)][String]$Title,
        [Parameter(mandatory=$true)][String]$Label
    )

    $Form = New-Object System.Windows.Forms.Form

    $Form.width = 300
    $Form.height = 150
    $Form.Text = $Title

    $DropDown = new-object System.Windows.Forms.ComboBox
    $DropDown.Location = new-object System.Drawing.Size(100,10)
    $DropDown.Size = new-object System.Drawing.Size(130,30)

    ForEach ($Item in $DropDownArray) {
     [void] $DropDown.Items.Add($Item)
    }

    $Form.Controls.Add($DropDown)

    $DropDownLabel = new-object System.Windows.Forms.Label
    $DropDownLabel.Location = new-object System.Drawing.Size(10,10) 
    $DropDownLabel.size = new-object System.Drawing.Size(100,40) 
    $DropDownLabel.Text = $Label
    $Form.Controls.Add($DropDownLabel)

    $Button = new-object System.Windows.Forms.Button
    $Button.Location = new-object System.Drawing.Size(100,50)
    $Button.Size = new-object System.Drawing.Size(100,20)
    $Button.Text = "Ok"
    $Button.Add_Click({Return-DropDown})
    $form.Controls.Add($Button)
    $form.ControlBox = $false

    $Form.Add_Shown({$Form.Activate()})
    [void] $Form.ShowDialog()


    return $script:choice
}

function Show-Window
{
[CmdletBinding(DefaultParameterSetName="OKWindow")]
Param (
    [Parameter(Mandatory=$true,
        ParameterSetName="OKWindow")]
    [Parameter(Mandatory=$true,
        ParameterSetName="OKCancelWindow")]
    [Parameter(Mandatory=$true,
        ParameterSetName="YesNoWindow")]
    [Parameter(Mandatory=$true,
        ParameterSetName="PrintWindow")]
    [String]$Title,
    [Parameter(Mandatory=$false,
        ParameterSetName="OKWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="OKCancelWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="YesNoWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="PrintWindow")]
    [Switch]$AddVScrollBar,
    [Parameter(Mandatory=$false,
        ParameterSetName="OKWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="OKCancelWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="YesNoWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="PrintWindow")]
    [string]$Comment,
    [Parameter(Mandatory=$true,
        ParameterSetName="OKWindow")]
    [Parameter(Mandatory=$true,
        ParameterSetName="OKCancelWindow")]
    [Parameter(Mandatory=$true,
        ParameterSetName="YesNoWindow")]
    [Parameter(Mandatory=$true,
        ParameterSetName="PrintWindow")]
    $Text,
    [Parameter(Mandatory=$false,
        ParameterSetName="OKWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="OKCancelWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="YesNoWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="PrintWindow")]
    [int32]$width,
    [Parameter(Mandatory=$false,
        ParameterSetName="OKWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="OKCancelWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="YesNoWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="PrintWindow")]
    [int32]$height,
    [Parameter(Mandatory=$false,
        ParameterSetName="OKWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="OKCancelWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="YesNoWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="PrintWindow")]
    [boolean]$AlwaysTop,
    [Parameter(Mandatory=$false,
        ParameterSetName="OKWindow")]
    [Switch]$OKWindow,
    [Parameter(Mandatory=$false,
        ParameterSetName="OKCancelWindow")]
    [Switch]$OKCancelWindow,
    [Parameter(Mandatory=$false,
        ParameterSetName="YesNoWindow")]
    [Switch]$YesNoWindow,
    [Parameter(Mandatory=$false,
        ParameterSetName="PrintWindow")]
    [Switch]$PrintWindow,
    [Parameter(Mandatory=$false,
        ParameterSetName="OKWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="OKCancelWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="YesNoWindow")]
    [Parameter(Mandatory=$false,
        ParameterSetName="PrintWindow")]
    [Switch]$ReadOnly
)

    if(!$width){$width=1000}
    if(!$height){$height=560}
    if(!$OKWindow -and !$OKCancelWindow -and !$YesNoWindow -and !$PrintWindow){$OKWindow=$true}

    if ($Text.GetType().Name -ne "String") {
    # we now assume that we got a text arrary instead of a string text
    # we need to convert this
        foreach ($line in $Text) {
            $NewTest += ($line+"`r`n")
        }
        $Text = $NewTest
        $NewTest = ""
    }

    $Script:BtnResult=$null
    $Comment=%{if($Comment){$Comment}else{" "}}

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = $Title
    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.AutoSize = $True
    $objTextBox = New-Object System.Windows.Forms.TextBox

    #region build UI
    $objForm.Topmost = %{if($AlwaysTop){$AlwaysTop}else{$False}}
    #$objForm.ControlBox = $false
    $objForm.FormBorderStyle = "FixedDialog"
    $objForm.StartPosition = "CenterScreen"
    $objForm.MinimizeBox = $False
    $objForm.MaximizeBox = $False
    $objForm.WindowState = "Normal"
    $objForm.Size = New-Object System.Drawing.Size($width,$height) 
    $objForm.BackColor = "White"
    $objForm.Icon = $Icon
    $objForm.Font = $FontStdt
    $objLabel.Location = New-Object System.Drawing.Size(10,10)
    $objLabel.Text = $Comment 
    $objTextBox.Location = New-Object System.Drawing.Size(15,50)
    $objTextBox.Size = New-Object System.Drawing.Size(($width-40),($height-160))
    $objTextBox.MultiLine = $True
    $objTextBox.ScrollBars = %{if($AddVScrollBar){"Vertical"}else{"None"}}
    $objTextBox.Font= New-Object System.Drawing.Font("Courier New",12,[System.Drawing.FontStyle]::Bold)
    $objTextBox.ForeColor = [System.Drawing.Color]::Green
    if ($ReadOnly) {
        $objTextBox.ReadOnly = $true
    }
    $objTextBox.Text=$Text
    $objTextBox.TabStop = $false

    $objForm.Controls.Add($objLabel)
    $objForm.Controls.Add($objTextBox)

    if($OKWindow){
        $objBtnOk = New-Object System.Windows.Forms.Button
        $objBtnOk.Cursor = [System.Windows.Forms.Cursors]::Hand
        #$objBtnOk.BackColor = [System.Drawing.Color]::LightGreen
        #$objBtnOk.Font = New-Object System.Drawing.Font("Verdana",14,,[System.Drawing.FontStyle]::Bold)
        $objBtnOk.Location = New-Object System.Drawing.Size((($width/2)-40),($height-100))
        $objBtnOk.Size = New-Object System.Drawing.Size(80,40)
        $objBtnOk.Text = "OK"
        $objBtnOk.Add_Click({
            $script:BtnResult="OK"
            $objForm.Close()
            $objForm.dispose()
        })
        $objForm.Controls.Add($objBtnOk)
    }elseif($OKCancelWindow){
        $objBtnOk = New-Object System.Windows.Forms.Button
        $objBtnOk.Cursor = [System.Windows.Forms.Cursors]::Hand
        #$objBtnOk.BackColor = [System.Drawing.Color]::LightGreen
        #$objBtnOk.Font = New-Object System.Drawing.Font("Verdana",14,,[System.Drawing.FontStyle]::Bold)
        $objBtnOk.Location = New-Object System.Drawing.Size((($width/4)-40),($height-100))
        $objBtnOk.Size = New-Object System.Drawing.Size(80,40)
        $objBtnOk.Text = "OK"
        $objBtnOk.Add_Click({
            $script:BtnResult="OK"
            $objForm.Close()
            $objForm.dispose()
        })
        $objBtnCancel = New-Object System.Windows.Forms.Button
        $objBtnCancel.Cursor = [System.Windows.Forms.Cursors]::Hand
        #$objBtnCancel.BackColor = [System.Drawing.Color]::LightGreen
        #$objBtnCancel.Font = New-Object System.Drawing.Font("Verdana",14,,[System.Drawing.FontStyle]::Bold)
        $objBtnCancel.Location = New-Object System.Drawing.Size(((($width/4)*3)-40),($height-100))
        $objBtnCancel.Size = New-Object System.Drawing.Size(80,40)
        $objBtnCancel.Text = "Cancel"
        $objBtnCancel.Add_Click({
            $script:BtnResult="CANCEL"
            $objForm.Close()
            $objForm.dispose()
        })
        $objBtnCancel.TabIndex=0
        $objForm.Controls.Add($objBtnOk)
        $objForm.Controls.Add($objBtnCancel)
    }elseif($YesNoWindow){
        $objBtnYes = New-Object System.Windows.Forms.Button
        $objBtnYes.Cursor = [System.Windows.Forms.Cursors]::Hand
        #$objBtnYes.BackColor = [System.Drawing.Color]::LightGreen
        #$objBtnYes.Font = New-Object System.Drawing.Font("Verdana",14,,[System.Drawing.FontStyle]::Bold)
        $objBtnYes.Location = New-Object System.Drawing.Size((($width/4)-40),($height-100))
        $objBtnYes.Size = New-Object System.Drawing.Size(80,40)
        $objBtnYes.Text = "Yes"
        $objBtnYes.Add_Click({
            $script:BtnResult="YES"
            $objForm.Close()
            $objForm.dispose()
        })
        $objBtnNo = New-Object System.Windows.Forms.Button
        $objBtnNo.Cursor = [System.Windows.Forms.Cursors]::Hand
        #$objBtnNo.BackColor = [System.Drawing.Color]::LightGreen
        #$objBtnNo.Font = New-Object System.Drawing.Font("Verdana",14,,[System.Drawing.FontStyle]::Bold)
        $objBtnNo.Location = New-Object System.Drawing.Size(((($width/4)*3)-40),($height-100))
        $objBtnNo.Size = New-Object System.Drawing.Size(80,40)
        $objBtnNo.Text = "No"
        $objBtnNo.Add_Click({
            $script:BtnResult="NO"
            $objForm.Close()
            $objForm.dispose()
        })
        $objBtnNo.TabIndex=0
        $objForm.Controls.Add($objBtnYes)
        $objForm.Controls.Add($objBtnNo)
    }else{ #$PrintWindow
        $objBtnOk = New-Object System.Windows.Forms.Button
        $objBtnOk.Cursor = [System.Windows.Forms.Cursors]::Hand
        #$objBtnOk.BackColor = [System.Drawing.Color]::LightGreen
        #$objBtnOk.Font = New-Object System.Drawing.Font("Verdana",14,,[System.Drawing.FontStyle]::Bold)
        $objBtnOk.Location = New-Object System.Drawing.Size((($width/4)-40),($height-100))
        $objBtnOk.Size = New-Object System.Drawing.Size(80,40)
        $objBtnOk.Text = "OK"
        $objBtnOk.Add_Click({
            $script:BtnResult="OK"
            $objForm.Close()
            $objForm.dispose()
        })
        $objBtnPrint = New-Object System.Windows.Forms.Button
        $objBtnPrint.Cursor = [System.Windows.Forms.Cursors]::Hand
        #$objBtnCancel.BackColor = [System.Drawing.Color]::LightGreen
        #$objBtnCancel.Font = New-Object System.Drawing.Font("Verdana",14,,[System.Drawing.FontStyle]::Bold)
        $objBtnPrint.Location = New-Object System.Drawing.Size(((($width/4)*3)-40),($height-100))
        $objBtnPrint.Size = New-Object System.Drawing.Size(80,40)
        $objBtnPrint.Text = "Print"
        $objBtnPrint.Add_Click({
            Print-Text $Text
            $script:BtnResult="OK"
            $objForm.Close()
            $objForm.dispose()
        })
        $objBtnPrint.TabIndex=0
        $objForm.Controls.Add($objBtnOk)
        $objForm.Controls.Add($objBtnPrint)
    }

    $objForm.Add_Shown({$objForm.Activate()})
    [void]$objForm.ShowDialog()
    #endregion
    return $BtnResult
}

function Submit-Request
{
    param(
        [Parameter(mandatory=$true)][String]$TargetCACnfg,
        [Parameter(mandatory=$true)][String]$TemplateName,
        [Parameter(mandatory=$true)][String]$RequestFile
    )

    
    
    $FolderName = $RequestFile.Substring(0,$RequestFile.Length - ((($RequestFile).split("\")[(($RequestFile).split("\")).Length-1]).Length))
    $CertFileName = $FolderName+((($RequestFile).split("\")[(($RequestFile).split("\")).Length-1]).split(".")[0])+".cer"
    $P7bFileName = $FolderName+((($RequestFile).split("\")[(($RequestFile).split("\")).Length-1]).split(".")[0])+".p7b"
    $MsgText = ""

#    $CertFileName = $FolderName + "CertNew-RequestID" + $RequestID + ".cer"
#    $P7BFileName = $FolderName + "CertNew-RequestID" + $RequestID + ".p7b"

    # submitting request
    $result = ""
    $result = certreq -f -config $TargetCACnfg -submit -attrib "certificatetemplate:$TemplateName" $RequestFile $CertFileName $P7BFileName
    if (($result -like "*error*") -or ($result -like "*fail*")) {
        $Msg = ""
        $Msg = ("Submission of "+ $RequestFile +" failed!`r`n" + $result + "`r`n`r`nAborting ...")
        Break-MessageBox $Msg
        exit
    } elseif ($result -like "*Certificate request is pending*")  {
        $ReqID = ($Result[0].split(":")[1]).trim()

    }

    # finally
    $MsgText = ("Submission of "+ $RequestFile +" succeeded!`r`n`r`n" + $result)
    #Show-Window -Title "Certificate request Submission Status" -Comment $TargetCACnfg -Text $MsgText -height 400 -width 700 -AddVScrollBar|Out-Null
    
    return $MsgText
}

Function Retrieve-Cert
{
    param(
        [Parameter(mandatory=$true)][String]$TargetCACnfg,
        [Parameter(mandatory=$true)][String]$RequestID,
        [Parameter(mandatory=$true)][String]$FolderName
    )

    $MsgText = ""

    # retrieving certificate
    $CertFileName = $FolderName + "CertNew-RequestID" + $RequestID + ".cer"
    $P7BFileName = $FolderName + "CertNew-RequestID" + $RequestID + ".p7b"

    $result = ""
    $result = Certreq -f -retrieve -config $TargetCACnfg $RequestID $CertFileName $P7BFileName
    if (($result -like "*error*") -or ($result -like "*fail*")) {
        $MsgText = ("Retrieval of certificate for RequestID "+ $RequestID +" failed!`r`n`r`n" + $result)
        return $MsgText
    }
    # removing rsp file
    Remove-Item ($FolderName + $CertFileName.split("\")[($CertFileName.split("\")).Length-1].split(".")[0]+".rsp")

    # finally
    $MsgText = ("Retrieval of certificate with RequestID "+ $RequestID + " into files:`r`n" + $CertFileName + "`r`n")
    $MsgText += ($P7BFileName +"`r`nsucceeded!`r`n`r`n" + $result)
    return $MsgText
}

function Get-IssuedCertList
{
    param(
        [Parameter(mandatory=$true)][String]$TargetCACnfg,
        [Parameter(mandatory=$false)][String]$RequesterName,
        [Parameter(mandatory=$true)][String]$TimeWindow
    )

    $objResultTextBox.Text = " "
    # defining target list with dimensions (chars): ReqID:10, RequesterName:40, CommonName:30, SAN:50
    $aResultList = @()
    $NotBefore = ((Get-Date).AddDays(-($TimeWindow))).ToShortDateString()
    if ($RequesterName) {
        $result = ConvertFrom-Csv (certutil -config $TargetCACnfg -view -restrict "disposition=20,RequesterName=$RequesterName,notbefore>=$NotBefore" -out "RequestID, Requestername, CommonName" csv)
    } else {
        $result = ConvertFrom-Csv (certutil -config $TargetCACnfg -view -restrict "disposition=20,notbefore>=$NotBefore" -out "RequestID, Requestername, CommonName" csv)
    }
    if ($result -ne $null) {
        if ($result -like "*fail*") {
            $objResultTextBox.Text = ("CA database query failed!`r`n`r`n" + $result.ToString())
        } else {
            $ListHeader = $result[0].psobject.properties.name #getting column names from PsObject (csv) as certutil will drop out localized names instead of real db schema names
            # crawling for SAN extensions
            $result | ForEach-Object {
                $objResult = "" | Select-Object RequestID,RequesterName,CommonName,SAN
                $objResult.RequestID = $_.($ListHeader[0])
                $objResult.RequesterName = $_.($ListHeader[1])
                $objResult.CommonName = $_.($ListHeader[2])
                $SanResult = ConvertFrom-Csv (certutil -config $TargetCACnfg -view -restrict "ExtensionRequestId=$($_.($ListHeader[0])), ExtensionName=2.5.29.17" ext csv)
                if ($SanResult) {
                    $SANListHeader = $SanResult[0].psobject.properties.name #getting column names from PsObject (csv) as certutil will drop out localized names instead of real db schema names
                    $i = 1
                    do {
                        if (($SanResult[$i].($SANListHeader[0]) -notlike "Other Name*")) {
                            if ($objResult.SAN) {
                                $objResult.SAN += ("|"+$SanResult[$i].($SANListHeader[0]))
                            } else {
                                $objResult.SAN += $SanResult[$i].($SANListHeader[0])
                            }
                        }
                    } while ($i++ -lt $SanResult.count)
                } else {
                    $sResult += "|"
                    $objResult.SAN = ""
                }
                $aResultList += $objResult
            }
        }
    }
    return $aResultList
}

Function Retrieve-PendingCerts
{
    param(
        [Parameter(mandatory=$true)][String]$TargetCACnfg
    )

    $TmpFileName = $env:temp + "\TmpDump.tmp"

    $objResultTextBox.Text = " "
    # defining target list with dimensions (chars): ReqID:10, RequesterName:40, CommonName:30, SAN:50
    $aResultList = @()

    # verify if request has been issued
    $result = ConvertFrom-Csv (certutil -config $TargetCACnfg -view -restrict "disposition=9" -out "RequestID, Requestername, CommonName, RawRequest" csv)
    if ($result -ne $null) {
        if ($result -like "*fail*") {
            $objResultTextBox.Text = ("CA database query failed!`r`n`r`n" + $result.ToString())
        } else {
            $ListHeader = $result[0].psobject.properties.name #getting column names from PsObject (csv) as certutil will drop out localized names instead of real db schema names
            # crawling for SAN extensions
            $result | ForEach-Object {
                $RawResult = certutil -config $TargetCACnfg -view -restrict "requestid=$($_.($ListHeader[0]))" -out Request.RawRequest > $TmpFileName
                if ($RawResult -match "Binary Request: EMPTY") {
                    $Msg = "Could not load certificate request on request ID " + $_.($ListHeader[0]) + "..."
                } elseif ($RawResult -match "Too many arguments") {
                    $Msg = "Error in query syntax!`r`n" + $RawResult 
                } elseif ($RawResult -match "CertUtil [Options]") {
                    $Msg = "Error in query syntax!`r`n" + $RawResult 
                } else {
                    $RawResult|Set-Content $TmpFileName -ErrorAction SilentlyContinue
                    if (Test-Path $TmpFileName -ErrorAction Ignore) {
                        $CertDump =Dump-Request -ReqFileName $TmpFileName
                        if ($CertDump -match "command FAILED") {
                            $Msg = "Could not enumerate SANs on request ID " + $_.($ListHeader[0]) + "..."
                        } else {
                            Remove-Item $TmpFileName -Force
                            $objResult = "" | Select-Object RequestID,RequesterName,CommonName,SAN
                            $objResult.RequestID = $_.($ListHeader[0])
                            $objResult.RequesterName = $_.($ListHeader[1])
                            $objResult.CommonName = $_.($ListHeader[2])
                            $CertDump.SAN|ForEach-Object {
                                $objResult.SAN = $objResult.SAN + ";" + $_
                            }
                        }
                    }
                    $aResultList += $objResult
                }
            }
        }
    }
    return $aResultList
}

function Refresh-PendingList
{
    param(
        [Parameter(mandatory=$true)][String]$TargetCACnfg
    )

    $objCertRetrieveListBox.Rows.Clear()
    (Retrieve-PendingCerts -TargetCACnfg $Script:CACnfg) | ForEach-Object {
        $objCertRetrieveListBox.Rows.Add($_.RequestID,$_.RequesterName, $_.CommonName,$_.SAN)
    }

}

function Display-RawData
{
    param(
        [Parameter(mandatory=$true)][String]$TargetCACnfg,
        [Parameter(mandatory=$true)][String]$Mode,
        [Parameter(mandatory=$false)][String]$RequestID
    )
    
    $TmpFileName = $env:temp + "\TmpDump.tmp"
    $Msg = ""
    $result = ""
    $CertDump = ""

    switch ($Mode) {
        "Retrieval" {
            $result = certutil -config $TargetCACnfg -view -restrict "requestid=$RequestID" -out rawcertificate
            if ($result -match "Binary Certificate: EMPTY") {
                $Msg = "Could not load certificate on request ID " + $RequestID + "..."
            } elseif ($result -match "Too many arguments") {
                $Msg = "Error in query syntax!`r`n" + $result 
            } elseif ($result -match "CertUtil [Options]") {
                $Msg = "Error in query syntax!`r`n" + $result 
            } else {
                $result|Set-Content $TmpFileName -ErrorAction SilentlyContinue
                if (Test-Path $TmpFileName) {
                    $CertDump = Certutil $TmpFileName
                    if ($CertDump -match "command FAILED") {
                        $Msg = "Could not dump certificate on request ID " + $RequestID + "..."
                    } else {
                        Show-Window -Title "Raw Certificate Data" -AddVScrollBar -OKWindow -Text $CertDump -width 1000 -height 800 -AlwaysTop $true -ReadOnly
                        $Msg = "Operation completed successful ..."
                    }
                    Remove-Item $TmpFileName -Force | Out-Null
                } else {
                    $Msg = "Could not dump certificate on request ID " + $RequestID + "..."
                }
            }
        }
        "ShowPending" {
            $result = certutil -config $TargetCACnfg -view -restrict "requestid=$RequestID" -out Request.RawRequest
            if ($result -match "Binary Request: EMPTY") {
                $Msg = "Could not load certificate request on request ID " + $RequestID + "..."
            } elseif ($result -match "Too many arguments") {
                $Msg = "Error in query syntax!`r`n" + $result 
            } elseif ($result -match "CertUtil [Options]") {
                $Msg = "Error in query syntax!`r`n" + $result 
            } else {
                $result|Set-Content $TmpFileName -ErrorAction SilentlyContinue
                if (Test-Path $TmpFileName) {
                    $CertDump = Certutil $TmpFileName
                    if ($CertDump -match "command FAILED") {
                        $Msg = "Could not dump certificate request on request ID " + $RequestID + "..."
                    } else {
                        Show-Window -Title "Raw Certificate Request Data" -AddVScrollBar -OKWindow -Text $CertDump -width 1000 -height 800 -AlwaysTop $true -ReadOnly
                        $Msg = "Operation completed successful ..."
                    }
                    Remove-Item $TmpFileName -Force
                } else {
                    $Msg = "Could not dump certificate request on request ID " + $RequestID + "..."
                }
            }
        }
    }
    return $Msg
}

function Issue-CertRequest
{
    param(
        [Parameter(mandatory=$true)][String]$TargetCACnfg,
        [Parameter(mandatory=$false)][String]$RequestID
    )

    $Msg = ""
    $result = ""

    $result = certutil -config $TargetCACnfg -resubmit $RequestID
    if ($result -match "FAILED") {
        # resubmission failed for some reason
        $Msg = "Approval error for request "+$RequestID+"!`r`nError Message:`r`n"+$result
        if ($result -like "*Access is denied*") {
            $Msg += "`r`n`r`nEnsure your account has 'Issue and Manage Certificates' permissions!"
        }
    } else {
        $Msg = "Request "+$RequestID+" successfully issued"
        Refresh-PendingList -TargetCACnfg $TargetCACnfg
    }
    Out-Message $Msg

}

function Deny-CertRequest
{
    param(
        [Parameter(mandatory=$true)][String]$TargetCACnfg,
        [Parameter(mandatory=$false)][String]$RequestID
    )

    $Msg = ""
    $result = ""

    $result = certutil -config $TargetCACnfg -deny $RequestID
    if ($result -match "FAILED") {
        # resubmission failed for some reason
        $Msg = "Could not process denial for request "+$RequestID+"!`r`nError Message:`r`n"+$result
        if ($result -like "*Access is denied*") {
            $Msg += "`r`n`r`nEnsure your account has 'Issue and Manage Certificates' permissions!"
        }
    } else {
        $Msg = "Request "+$RequestID+" successfully denied"
        Refresh-PendingList -TargetCACnfg $TargetCACnfg
    }
    Out-Message $Msg


}
