<#
 ==============================================================================================
 THIS SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
 OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
 FITNESS FOR A PARTICULAR PURPOSE.
 ==============================================================================================
 #>

 <#

    .SYNOPSIS
    Submitts a request to the CA or retrieves an issued certificate from a CA.
    
    .PARAMETER Help
    display help.

   .Notes
   Version 1.1

#>



$BaseDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# loading all Libary Scripts we depend on
Get-ChildItem -Path "$Script:BaseDirectory" -Filter lib*.ps1 | ForEach-Object {
    Try {
        Write-Host "[INFO] Loading library $($_.FullName)" -ForegroundColor Green
        . ($_.FullName)
    }
    Catch {
        Write-host "Error loading $($_.FullName)" -ForegroundColor Yellow -BackgroundColor DarkRed
        Exit
    }
}

If ($Help) {
    Get-Help $MyInvocation.MyCommand.Definition -Detailed
    Exit
}

#get current logged on user name
$MyIdentity = ([system.security.principal.windowsidentity]::GetCurrent()).Name

#some necessary variables and initializations
$Script:CACnfg = ""          # CA config string
$Script:CAOpBlocked = $False # [boolean] indicates if CA can be addressed
$Script:CertTask = "Submit"  # [string] defines operation mode (Submit|Retrieval|ShowPending)
$Script:intSelectedTime = 1  # numeric representation of the selected time interval from array below
$aTimeInterval = ("1 day","7 days","14 days","30 days","90 days")

# collecting enrollment CAs from AD --> certutil -adca
$aEntCAs = Get-AdEnrollCaList
if (!($aEntCAs)) {
    # no CAs found - that's weird and should be investigated
    Break-MessageBox -Message "No Enrollment-CAs found!`r`n`r`nAborting ..."
    Exit
}

#region build UI
#form and form panel dimensions
    $width = 600
    $height = 650
    $Panelwidth = $Width-30
    $Panelheight = $Height-340

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Certificate Management"

    $objForm.Topmost = $False
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

    $objCADropDownBoxLabel = new-object System.Windows.Forms.Label
    $objCADropDownBoxLabel.Location = new-object System.Drawing.Point(10,10) 
    $objCADropDownBoxLabel.size = new-object System.Drawing.Size(160,40) 
    $objCADropDownBoxLabel.Text = "Select Target CA:"

    $objCADropDownBox = new-object System.Windows.Forms.ComboBox
    $objCADropDownBox.Location = new-object System.Drawing.Point(195,10)
    $objCADropDownBox.Size = new-object System.Drawing.Size(330,30)

#form function radio buttons
    $flp = New-Object System.Windows.Forms.FlowLayoutPanel
    $flpwidth = $objForm.Width.ToString()
    $flp.Size = "$flpwidth,30"
    $flp.Location = New-Object System.Drawing.Point(10,50)
    $flp.FlowDirection = 'LeftToRight'

    $rb1 = New-Object System.Windows.Forms.RadioButton
    $rb1.Text = "Submit Request"
    $rb1.AutoSize = $true
    $rb1.Checked = $true

    $rb2 = New-Object System.Windows.Forms.RadioButton
    $rb2.Text = "Retrieve Certificate"
    $rb2.AutoSize = $true

    $flp.Controls.Add($rb1)
    $flp.Controls.Add($rb2)

#region SubmitPanel

    $objSubmitPanel = New-Object System.Windows.Forms.Panel
    $objSubmitPanel.Location = new-object System.Drawing.Point(10,90)
    $objSubmitPanel.size = new-object System.Drawing.Size($Panelwidth,$Panelheight) 
    #$objSubmitPanel.BackColor = "255,0,255"
    #$objSubmitPanel.BackColor = "Blue"
    $objSubmitPanel.BorderStyle = "FixedSingle"

    $objCsrInputLabel = new-object System.Windows.Forms.Label
    $objCsrInputLabel.Location = new-object System.Drawing.Point(10,10) 
    $objCsrInputLabel.size = new-object System.Drawing.Size(170,30) 
    $objCsrInputLabel.Text = "Select request file:"


#file select input
    $objCsrInputField = New-Object System.Windows.Forms.TextBox
    $objCsrInputField.Location = New-Object System.Drawing.Point(185,10)
    $objCsrInputField.Size = New-Object System.Drawing.Size(275,25)
    $objCsrInputField.Multiline = $false ### Allows multiple lines of data
    $objCsrInputField.AcceptsReturn = $false ### By hitting enter it creates a new line

    $objCsrFileSelectBtn = New-Object System.Windows.Forms.Button
    $objCsrFileSelectBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $objCsrFileSelectBtn.Location = New-Object System.Drawing.Point(470,10)
    $objCsrFileSelectBtn.Size = New-Object System.Drawing.Size(80,25)
    $objCsrFileSelectBtn.Text = "Browse"

    $objTmplDropDownBoxLabel = new-object System.Windows.Forms.Label
    $objTmplDropDownBoxLabel.Location = new-object System.Drawing.Point(10,50) 
    $objTmplDropDownBoxLabel.size = new-object System.Drawing.Size(170,40) 
    $objTmplDropDownBoxLabel.Text = "Select Template:"

    $objTmplDropDownBox = new-object System.Windows.Forms.ComboBox
    $objTmplDropDownBox.Location = new-object System.Drawing.Point(185,50)
    $objTmplDropDownBox.Size = new-object System.Drawing.Size(330,30)

    $objSubmitCsrTextBoxLabel = new-object System.Windows.Forms.Label
    $objSubmitCsrTextBoxLabel.Location = new-object System.Drawing.Point(10,90) 
    $objSubmitCsrTextBoxLabel.size = new-object System.Drawing.Size(($Panelwidth-20),25) 
    $objSubmitCsrTextBoxLabel.Text = "Request data:"

    $objSubmitCsrTextBox = New-Object System.Windows.Forms.TextBox
    $objSubmitCsrTextBox.Location = New-Object System.Drawing.Point(10,115)
    $objSubmitCsrTextBox.Size = New-Object System.Drawing.Size(($Panelwidth-20),($Panelheight-130))
    $objSubmitCsrTextBox.ReadOnly = $true 
    $objSubmitCsrTextBox.Multiline = $true
    $objSubmitCsrTextBox.AcceptsReturn = $true 

#finally build panel
    $objSubmitPanel.Controls.Add($objCsrInputLabel)
    $objSubmitPanel.Controls.Add($objCsrInputField)
    $objSubmitPanel.Controls.Add($objCsrFileSelectBtn)
    $objSubmitPanel.Controls.Add($objTmplDropDownBoxLabel)
    $objSubmitPanel.Controls.Add($objTmplDropDownBox)
    $objSubmitPanel.Controls.Add($objSubmitCsrTextBoxLabel)
    $objSubmitPanel.Controls.Add($objSubmitCsrTextBox)

#endregion


#region objRetrievePanel

    $objRetrievePanel = New-Object System.Windows.Forms.Panel
    $objRetrievePanel.Location = new-object System.Drawing.Point(10,90)
    $objRetrievePanel.size = new-object System.Drawing.Size($Panelwidth,$Panelheight) 
    #$objRetrievePanel.BackColor = "255,0,255"
    #$objRetrievePanel.BackColor = "Blue"
    $objRetrievePanel.BorderStyle = "FixedSingle"

#region context menu
    $contextMenuRetrieveStrip1 = New-Object System.Windows.Forms.ContextMenuStrip
    [System.Windows.Forms.ToolStripItem]$CxtRetrieveMnuStrip1Item1 = New-Object System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripItem]$CxtRetrieveMnuStrip1Item2 = New-Object System.Windows.Forms.ToolStripMenuItem
    [System.Windows.Forms.ToolStripItem]$CxtRetrieveMnuStrip1Item3 = New-Object System.Windows.Forms.ToolStripMenuItem
    $CxtRetrieveMnuStrip1Item1.Text = "Show Binary Data"
    $CxtRetrieveMnuStrip1Item2.Text = "Issue"
    $CxtRetrieveMnuStrip1Item3.Text = "Deny"
    $CxtRetrieveMnuStrip1Item2.Enabled = $False
    $CxtRetrieveMnuStrip1Item3.Enabled = $False
    $contextMenuRetrieveStrip1.Items.Add($CxtRetrieveMnuStrip1Item1)|Out-Null
    $contextMenuRetrieveStrip1.Items.Add($CxtRetrieveMnuStrip1Item2)|Out-Null
    $contextMenuRetrieveStrip1.Items.Add($CxtRetrieveMnuStrip1Item3)|Out-Null
#endregion

    $objTimeIntLabel = new-object System.Windows.Forms.Label
    $objTimeIntLabel.Location = new-object System.Drawing.Point(10,10) 
    $objTimeIntLabel.size = new-object System.Drawing.Size(175,30) 
    $objTimeIntLabel.Text = "Select time interval:"

    $objTimeIntDropDownBox = new-object System.Windows.Forms.ComboBox
    $objTimeIntDropDownBox.Location = new-object System.Drawing.Point(185,10)
    $objTimeIntDropDownBox.Size = new-object System.Drawing.Size(330,30)

# sub function radio buttons
    $objRetrieveRBMyReq = New-Object System.Windows.Forms.RadioButton
    $objRetrieveRBMyReq.Location = new-object System.Drawing.Point(10,40)
    $objRetrieveRBMyReq.Text = "Show my requests"
    $objRetrieveRBMyReq.AutoSize = $true
    $objRetrieveRBMyReq.Checked = $true

    $objRetrieveRBAllReq = New-Object System.Windows.Forms.RadioButton
    $objRetrieveRBAllReq.Location = new-object System.Drawing.Point(200,40)
    $objRetrieveRBAllReq.Text = "Show all requests"
    $objRetrieveRBAllReq.AutoSize = $true

    $objShowPendingRB = New-Object System.Windows.Forms.RadioButton
    $objShowPendingRB.Location = new-object System.Drawing.Point(400,40)
    $objShowPendingRB.Text = "Show pending"
    $objShowPendingRB.AutoSize = $true
    $objShowPendingRB.Checked = $false

# database extract view
    $objCertRetrieveListBox = New-Object System.Windows.Forms.DataGridView 
    $objCertRetrieveListBox.Location = New-Object System.Drawing.Size(10,80) 
    $objCertRetrieveListBox.Size = New-Object System.Drawing.Size(($Panelwidth-20),180)
    $objCertRetrieveListBox.DefaultCellStyle.Font = "Microsoft Sans Serif, 9"
    $objCertRetrieveListBox.ColumnHeadersDefaultCellStyle.Font = "Microsoft Sans Serif, 9"
    $objCertRetrieveListBox.ColumnCount = 4
    $objCertRetrieveListBox.ColumnHeadersVisible = $true
    $objCertRetrieveListBox.SelectionMode = "FullRowSelect"
    $objCertRetrieveListBox.ReadOnly = $true
    $objCertRetrieveListBox.Columns[0].Name = "RequestID"
    $objCertRetrieveListBox.Columns[1].Name = "RequesterName"
    $objCertRetrieveListBox.Columns[2].Name = "CommonName"
    $objCertRetrieveListBox.Columns[3].Name = "SAN"

    $objCertRetrieveListBox.Columns[0].Width = 80
    $objCertRetrieveListBox.Columns[1].Width = 320
    $objCertRetrieveListBox.Columns[2].Width = 200
    $objCertRetrieveListBox.Columns[3].Width = 500
    $objCertRetrieveListBox.ContextMenuStrip = $contextMenuRetrieveStrip1

    $objCertInputLabel = new-object System.Windows.Forms.Label
    $objCertInputLabel.Location = new-object System.Drawing.Point(10,($Panelheight-50)) 
    $objCertInputLabel.size = new-object System.Drawing.Size(280,30) 
    $objCertInputLabel.Text = "Select certificate output folder:"

#file select output folder
    $objCertOutFolder = New-Object System.Windows.Forms.TextBox
    $objCertOutFolder.Location = New-Object System.Drawing.Point(10,($Panelheight-30))
    $objCertOutFolder.Size = New-Object System.Drawing.Size(275,25)
    $objCertOutFolder.Multiline = $false ### Allows multiple lines of data
    $objCertOutFolder.AcceptsReturn = $false ### By hitting enter it creates a new line
    $objCertOutFolder.ReadOnly = $true


    $objCertFolderSelectBtn = New-Object System.Windows.Forms.Button
    $objCertFolderSelectBtn.Cursor = [System.Windows.Forms.Cursors]::Hand
    $objCertFolderSelectBtn.Location = New-Object System.Drawing.Point(300,($Panelheight-30))
    $objCertFolderSelectBtn.Size = New-Object System.Drawing.Size(80,25)
    $objCertFolderSelectBtn.Text = "Browse"

#finally build panel
    $objRetrievePanel.Controls.Add($objTimeIntLabel)
    $objRetrievePanel.Controls.Add($objTimeIntDropDownBox)
    $objRetrievePanel.Controls.Add($objCertRetrieveListBox)
    $objRetrievePanel.Controls.Add($objRetrieveRBMyReq)
    $objRetrievePanel.Controls.Add($objRetrieveRBAllReq)
    $objRetrievePanel.Controls.Add($objShowPendingRB)
    $objRetrievePanel.Controls.Add($objCertOutFolder)
    $objRetrievePanel.Controls.Add($objCertInputLabel)
    $objRetrievePanel.Controls.Add($objCertFolderSelectBtn)

#endregion

#Operation result text box
    $objResultTextBoxLabel = new-object System.Windows.Forms.Label
    $objResultTextBoxLabel.Location = new-object System.Drawing.Point(10,($height-245)) 
    $objResultTextBoxLabel.size = new-object System.Drawing.Size(130,25) 
    $objResultTextBoxLabel.Text = "Output log:"

    $objResultTextBox = New-Object System.Windows.Forms.TextBox
    $objResultTextBox.Location = New-Object System.Drawing.Point(10,($height-220))
    $objResultTextBox.Size = New-Object System.Drawing.Size(($width-30),130)
    $objResultTextBox.ReadOnly = $true 
    $objResultTextBox.Multiline = $true
    $objResultTextBox.AcceptsReturn = $true 
    $objResultTextBox.Text = ""

# form execution buttons
    $objBtnOk = New-Object System.Windows.Forms.Button
    $objBtnOk.Cursor = [System.Windows.Forms.Cursors]::Hand
    $objBtnOk.Location = New-Object System.Drawing.Point((($width/4)-40),($height-80))
    $objBtnOk.Size = New-Object System.Drawing.Size(90,30)
    $objBtnOk.Text = "Submit"

    $objBtnExit = New-Object System.Windows.Forms.Button
    $objBtnExit.Cursor = [System.Windows.Forms.Cursors]::Hand
    $objBtnExit.Location = New-Object System.Drawing.Point(((($width/4)*3)-40),($height-80))
    $objBtnExit.Size = New-Object System.Drawing.Size(80,30)
    $objBtnExit.Text = "Exit"
    $objBtnExit.TabIndex=0

# complete form construct    
    $objForm.Controls.Add($objCADropDownBoxLabel)
    $objForm.Controls.Add($objCADropDownBox)
    $objForm.Controls.Add($flp)
    $objForm.Controls.Add($objSubmitPanel)
    $objForm.Controls.Add($objResultTextBoxLabel)
    $objForm.Controls.Add($objRetrievePanel)
    $objForm.Controls.Add($objResultTextBox)
    $objForm.Controls.Add($objBtnOk)
    $objForm.Controls.Add($objBtnExit)
#endregion

#region form event handlers
    $objCADropDownBox.Add_SelectedValueChanged({
        ReLoad-FormContent -arrEntCAs $aEntCAs -CaDropDownCurrSelectionIndex $objCADropDownBox.SelectedIndex
    })

    $rb1.Add_Click({
        $Script:CertTask = "Submit"
        $objRetrievePanel.Hide()
        $objSubmitPanel.Show()
        $objBtnOk.Text = "Submit"
        $objResultTextBox.Text = ""
        if ($Script:CAOpBlocked) {$objBtnOk.Hide()}else{$objBtnOk.Show()}
    })

    $rb2.Add_Click({
        $Script:CertTask = "Retrieval"
        $objSubmitPanel.Hide()
        $objRetrievePanel.Show()
        $objBtnOk.Text = "Retrieve"
        $objResultTextBox.Text = ""
        if ($objShowPendingRB.Checked -or $Script:CAOpBlocked) {$objBtnOk.Hide()}else{$objBtnOk.Show()}
    })

    $objCsrInputField.Add_TextChanged({
        if ($objCsrInputField.Text) {
            if (Test-Path $objCsrInputField.Text) {
                $Csr = Dump-Request -ReqFileName $objCsrInputField.Text
                $objSubmitCsrTextBox.Text = ("Subject: " + $Csr.subjectName + "`r`nEnhanced Key Usage: " + $csr.enhancedKeyUsage + "`r`nKey Length: " + $Csr.KeyLength + "`r`n")
                if ($Csr.san) {
                    $objSubmitCsrTextBox.Text += ("SAN(s): " + ($Csr.san|ForEach-Object {("`r`n    " + $_.Type + " = " + $_.SAN)}))
                } else {
                    $objSubmitCsrTextBox.Text += ("SAN(s): " + "No SANs found in CSR!")
                }
            }
        }
    })

    $objCsrFileSelectBtn.Add_Click({
        $RequestFile = Select-File -StartFolder $BaseDirectory -FileFilter $CsrFilter
        if ($RequestFile) {
            $objCsrInputField.Text = $RequestFile
        }
    })

    $ClickElementMenu=
    {
        [System.Windows.Forms.ToolStripItem]$sender = $args[0]
        [System.EventArgs]$e= $args[1]

        $ReId = $objCertRetrieveListBox.CurrentRow.Cells.value[0]
        Switch ($sender.Text) {
            "Show Binary Data" {
                #Write-Host "Show Binary Data"
                Display-RawData -TargetCACnfg $Script:CaCnfg -Mode $Script:CertTask -RequestID ($objCertRetrieveListBox.CurrentRow.Cells.value[0])
            }
            "Issue" {
                Issue-CertRequest -TargetCACnfg $Script:CaCnfg -RequestID ($objCertRetrieveListBox.CurrentRow.Cells.value[0])
            }
            "Deny" {
                Deny-CertRequest -TargetCACnfg $Script:CaCnfg -RequestID ($objCertRetrieveListBox.CurrentRow.Cells.value[0])
            }
        }
    
    }
    
    $CxtRetrieveMnuStrip1Item1.add_Click($ClickElementMenu)
    $CxtRetrieveMnuStrip1Item2.add_Click($ClickElementMenu)
    $CxtRetrieveMnuStrip1Item3.add_Click($ClickElementMenu)

    $objTimeIntDropDownBox.Add_SelectedValueChanged({
        Switch ($objTimeIntDropDownBox.SelectedItem) {
            "1 day" {
                $Script:intSelectedTime = 1
            }
            "7 days" {
                $Script:intSelectedTime = 7
            }
            "14 days" {
                $Script:intSelectedTime = 14
            }
            "30 days" {
                $Script:intSelectedTime = 30
            }
            "90 days" {
                $Script:intSelectedTime = 90
            }
        }
        if (!$Script:CAOpBlocked) {
            $objCertRetrieveListBox.Rows.Clear()
            (Get-IssuedCertList -TargetCACnfg $Script:CACnfg -RequesterName $MyIdentity -TimeWindow $Script:intSelectedTime) | ForEach-Object {
                $objCertRetrieveListBox.Rows.Add($_.RequestID,$_.RequesterName, $_.CommonName,$_.SAN)
            }
        }
    })

    $objRetrieveRBMyReq.Add_Click({
        $Script:CertTask = "Retrieval"
        $CertReqs = "OnlyMy"

        #disable Issue|Deny context menue
        $CxtRetrieveMnuStrip1Item2.Enabled = $False
        $CxtRetrieveMnuStrip1Item3.Enabled = $False
        
        $objCertFolderSelectBtn.Enabled = $True
        $objCertOutFolder.Enabled = $True
        $objTimeIntDropDownBox.Enabled = $True
        $objBtnOk.Show()

        $objCertRetrieveListBox.Rows.Clear()
        (Get-IssuedCertList -TargetCACnfg $Script:CACnfg -RequesterName $MyIdentity -TimeWindow $Script:intSelectedTime) | ForEach-Object {
            $objCertRetrieveListBox.Rows.Add($_.RequestID,$_.RequesterName, $_.CommonName,$_.SAN)
        }
    })

    $objRetrieveRBAllReq.Add_Click({
        $Script:CertTask = "Retrieval"
        $CertReqs = "All"

        #disable Issue|Deny context menue
        $CxtRetrieveMnuStrip1Item2.Enabled = $False
        $CxtRetrieveMnuStrip1Item3.Enabled = $False

        $objCertFolderSelectBtn.Enabled = $True
        $objCertOutFolder.Enabled = $True
        $objTimeIntDropDownBox.Enabled = $True
        $objBtnOk.Show()

        $objCertRetrieveListBox.Rows.Clear()
        (Get-IssuedCertList -TargetCACnfg $Script:CACnfg -TimeWindow $Script:intSelectedTime) | ForEach-Object {
            $objCertRetrieveListBox.Rows.Add($_.RequestID,$_.RequesterName, $_.CommonName,$_.SAN)
        }
    })

    $objShowPendingRB.Add_Click({
        $Script:CertTask = "ShowPending"

        #disable Issue|Deny context menue
        $CxtRetrieveMnuStrip1Item2.Enabled = $True
        $CxtRetrieveMnuStrip1Item3.Enabled = $True

        $objCertFolderSelectBtn.Enabled = $False
        $objCertOutFolder.Enabled = $False
        $objTimeIntDropDownBox.Enabled = $False
        $objBtnOk.Hide()

        $objCertRetrieveListBox.Rows.Clear()
        (Retrieve-PendingCerts -TargetCACnfg $Script:CACnfg) | ForEach-Object {
            $objCertRetrieveListBox.Rows.Add($_.RequestID,$_.RequesterName, $_.CommonName,$_.SAN)
        }
    })

    $objCertRetrieveListBox.add_MouseDown({
        $sender = $args[0]
        [System.Windows.Forms.MouseEventArgs]$e= $args[1]

        if ($e.Button -eq  [System.Windows.Forms.MouseButtons]::Right)
        {
            [System.Windows.Forms.DataGridView+HitTestInfo] $hit = $objCertRetrieveListBox.HitTest($e.X, $e.Y);
            if ($hit.Type -eq [System.Windows.Forms.DataGridViewHitTestType]::Cell)
            {
                $objCertRetrieveListBox.CurrentCell = $objCertRetrieveListBox[$hit.ColumnIndex, $hit.RowIndex];
                $contextMenuRetrieveStrip1.Show($objCertRetrieveListBox, $e.X, $e.Y);
            }

        }
    })

    $objCertFolderSelectBtn.Add_Click({
        $CertOutFolder = Select-Folder -StartFolder $BaseDirectory
        if ($CertOutFolder) {
            $objCertOutFolder.Text = ($CertOutFolder+"\CertNew-RequestID-<ReqID>.cer/p7b")
        }
    })

    $objBtnOk.Add_Click({
#        $objResultTextBox.Text = ""

        Switch ($Script:CertTask) {
            "Submit" {
                if (!($objCsrInputField.Text)) {
                    $objResultTextBox.Text = ("Please select a request file to submit.`r`nAborting ...")

                } else {
                    $objResultTextBox.Text = Submit-Request -TargetCACnfg $Script:CACnfg -RequestFile $objCsrInputField.Text -TemplateName $objTmplDropDownBox.SelectedItem
                    $objCsrInputField.Clear()
                }
            }
            "Retrieval" {
                if (!($objCertOutFolder.Text)) {
                    $objResultTextBox.Text = ("Please select an output folder for the retrieved vertificates.`r`nAborting ...")

                } else {
                    if ($objCertOutFolder.Text -match "CertNew-RequestID-<ReqID>.cer/p7b") {
                        $FolderName = $objCertOutFolder.Text.Substring(0,$objCertOutFolder.Text.Length - ((($objCertOutFolder.Text).split("\")[(($objCertOutFolder.Text).split("\")).Length-1]).Length))
                    } else {
                        $FolderName = $objCertOutFolder.Text
                    }
                    $ReqID = ($objCertRetrieveListBox.CurrentRow.Cells.Item(0).value).trim() 
                    $objResultTextBox.Text = Retrieve-Cert -TargetCACnfg $Script:CACnfg -FolderName $FolderName -RequestID $ReqID
                    $objCertOutFolder.Clear()
                }
            }
            "ShowPending" {
                #do nothing
            }
        }
    })

    $objBtnExit.Add_Click({
        $script:BtnResult="Exit"
        $objForm.Close()
        $objForm.dispose()
    })

#endregion


Load-InitialFormContent -arrEntCAs $aEntCAs -arrTime $aTimeInterval
#(Get-IssuedCertList -TargetCACnfg $Script:CACnfg -RequesterName $MyIdentity -TimeWindow $Script:intSelectedTime) | ForEach-Object {
#    $objCertRetrieveListBox.Rows.Add($_.RequestID,$_.RequesterName, $_.CommonName,$_.SAN)
#}


$objForm.Add_Shown({$objForm.Activate()})
[void]$objForm.ShowDialog()







