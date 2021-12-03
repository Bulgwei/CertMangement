# ==============================================================================================
# THIS SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
# OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
# FITNESS FOR A PARTICULAR PURPOSE.
#
# This sample is not supported under any Microsoft standard support program or service. 
# The script is provided AS IS without warranty of any kind. Microsoft further disclaims all
# implied warranties including, without limitation, any implied warranties of merchantability
# or of fitness for a particular purpose. The entire risk arising out of the use or performance
# of the sample and documentation remains with you. In no event shall Microsoft, its authors,
# or anyone else involved in the creation, production, or delivery of the script be liable for 
# any damages whatsoever (including, without limitation, damages for loss of business profits, 
# business interruption, loss of business information, or other pecuniary loss) arising out of 
# the use of or inability to use the sample or documentation, even if Microsoft has been advised 
# of the possibility of such damages.
# ==============================================================================================
#
# version 1.1
# dev'd by andreas.luy@microsoft.com
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





