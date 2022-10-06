# ==============================================================================================
# THIS SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
# OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
# FITNESS FOR A PARTICULAR PURPOSE.
# ==============================================================================================
#
# version 1.1
#


function Load-InitialFormContent
{
    param(
        [Parameter(mandatory=$True)][Array]$arrEntCAs,
        [Parameter(mandatory=$True)][Array]$arrTime
    )

    $Msg = " "
    $Script:CACnfg = $null
    $Script:CACnfg = ($arrEntCAs[0].DnsHostName+"\"+$arrEntCAs[0].CaName) #1st element

    ForEach ($Item in $arrEntCAs) {
        [void] $objCADropDownBox.Items.Add($Item.CaName)
    }
    ForEach ($Item in $arrTime) {
        [void] $objTimeIntDropDownBox.Items.Add($Item)
    }

    $objCADropDownBox.SelectedItem = $objCADropDownBox.Items[0]
    $objTimeIntDropDownBox.SelectedItem = $objTimeIntDropDownBox.Items[0]
}

function ReLoad-FormContent
{
    param(
        [Parameter(mandatory=$True)][Array]$arrEntCAs, 
        [Parameter(mandatory=$True)][int32]$CaDropDownCurrSelectionIndex #$objCADropDownBox.SelectedIndex
    )


    $Msg = " "
    $arrEntTmpls = @()
    $Script:CACnfg = $null
    $Script:CACnfg = ($arrEntCAs[$CaDropDownCurrSelectionIndex].DnsHostName+"\"+$arrEntCAs[$CaDropDownCurrSelectionIndex].CaName)
    $objTmplDropDownBox.Items.Clear()
    $objCertRetrieveListBox.Rows.Clear()

    if (!(IsCAResponding -TargetCACnfg $Script:CACnfg)) {
        $Msg = ("Target CA " + $arrEntCAs[$CaDropDownCurrSelectionIndex].CaName + " is not responding!`r`n")
        $Msg += "Verify that the CA and that the CA is online and reachable.`r`n`r`nAborting ..."
        [void] $objTmplDropDownBox.Items.Add("CA Not Responding!")
        $Script:CAOpBlocked = $True
    } elseif (!(HasCAAdmin -TargetCACnfg $Script:CACnfg)) {
        $Msg = "Target CA " + $arrEntCAs[$CaDropDownCurrSelectionIndex].CaName         $Msg += " is not responding to administrative requests!`r`nVerify that your account has administrative "        $Msg += "permissions on the CA and that the CA`r`nis online and reachable.`r`n`r`nAborting ..."
        [void] $objTmplDropDownBox.Items.Add("No Permissions on CA!")
        $Script:CAOpBlocked = $True
    } else {
        $objTmplDropDownBox.Items.Clear()
        Get-CATemplates -TargetCACnfg $Script:CACnfg|ForEach-Object {
            $arrEntTmpls += ,@($arrEntCAs[$CaDropDownCurrSelectionIndex].CaName,$_)
            [void] $objTmplDropDownBox.Items.Add($_)
        }

        $Script:CAOpBlocked = $False
        Switch ($Script:CertTask) {
            "Submit" {
                # do nothing
            }
            "Retrieval" {
                $objRetrievePanel.Enabled = $True
                $objSubmitPanel.Enabled = $True
                $objBtnOk.Show()

                (Get-IssuedCertList -TargetCACnfg $Script:CACnfg -RequesterName $MyIdentity -TimeWindow $Script:intSelectedTime) | ForEach-Object {
                    $objCertRetrieveListBox.Rows.Add($_.RequestID,$_.RequesterName, $_.CommonName,$_.SAN)
                }
            }
            "ShowPending" {
                $objRetrievePanel.Enabled = $True
                $objSubmitPanel.Enabled = $True
                $objBtnOk.Hide()

                (Retrieve-PendingCerts -TargetCACnfg $Script:CACnfg) | ForEach-Object {
                    $objCertRetrieveListBox.Rows.Add($_.RequestID,$_.RequesterName, $_.CommonName,$_.SAN)
                }
            }
        } 

    }

    if($Script:CAOpBlocked){
        $objRetrievePanel.Enabled = $False
        $objSubmitPanel.Enabled = $False
        $objBtnOk.Hide()
    
    }

    Out-Message $Msg
    $objTmplDropDownBox.SelectedItem = $objTmplDropDownBox.Items[0]
}

function Reload-DbView
{

}

function Out-Message
{
    param(
        [Parameter(mandatory=$True)][String]$Msg
    )
    $objResultTextBox.Text = $Msg
}





