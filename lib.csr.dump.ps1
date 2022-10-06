# ==============================================================================================
# THIS SAMPLE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED 
# OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR 
# FITNESS FOR A PARTICULAR PURPOSE.
# ==============================================================================================
#
# version 1.1
#
#

function IsNumeric($value) 
{
   return ($($value.Trim()) -match "^[-]?[0-9.]+$")
}


function Check-Signature
{
    param(
        [Parameter(mandatory=$true)] [Array]$ReqString
    )

    #
    # return data type: boolean
    #

    $line=""
    $found=$false

    foreach ($line in $ReqString){
        if($line -match "Signature matches Public Key"){$found=$true}
    }
    return $found
}


function Get-RequestAttributes
{
    param(
        [Parameter(mandatory=$true)] [Array]$ReqString
    )

    #
    # return data type: array of custom PS objects
    # objects contain of: AttributeName,AttribField1,AttribField2,AttribField3,AttribField4,AttribField5
    #

    $line=""
    $lineCount=0
    $found=$false
    $fieldNum=0
    $fieldLine=""
    $Attributes=@()
    $numCount=0
    $Count=0
 	$obj="" | Select AttributeName,AttribField1,AttribField2,AttribField3,AttribField4,AttribField5

    foreach ($line in $ReqString){
        $lineCount++
        if($found){
            if($line -match "  Attribute"){
	            $obj="" | Select AttributeName,AttribField1,AttribField2,AttribField3,AttribField4,AttribField5
                $obj.AttributeName=($line.Split("(")[1]).Substring(0,($line.Split("(")[1]).length-1)
                $Count++
                $fieldNum=0
                $i=0
                do{
                    $fieldLine=$ReqString[$lineCount+$i++]
                    if(($fieldLine -notmatch "Value") -and ($fieldLine -notmatch "Unknown Attribute")){
                        switch(++$fieldNum)
                        {
                            1 {$obj.AttribField1=$fieldLine.trim()}
                            2 {$obj.AttribField2=$fieldLine.trim()}
                            3 {$obj.AttribField3=$fieldLine.trim()}
                            4 {$obj.AttribField4=$fieldLine.trim()}
                            5 {$obj.AttribField5=$fieldLine.trim()}
                        }
                    }
                }until(($ReqString[$lineCount+$i].length -eq 0) -or !($ReqString[$lineCount+$i].StartsWith(" ")))
                $Attributes += $obj

            }
        }
        if($line -match "Request Attributes:"){
            [int32]$numCount=$line.Split(":")[1]
            if($numCount -gt 0){$found=$true}
        }
        if($found -and ($Count -eq $numCount)){$found=$false}
    }
    return $Attributes
}

function Get-CertExtensions
{
    param(
        [Parameter(mandatory=$true)] [Array]$ReqString
    )
    #
    # return data type: array of custom PS objects
    # objects contain of: ExtensionOID,ExtensionName,ExtensionField
    #

    $line=""
    $lineCount=0
    $found=$false
    $fieldNum=0
    $fieldLine=""
    $Extensions=@()
    $numCount=0
    $Count=0
 	$obj="" | Select ExtensionOID,ExtensionName,ExtensionField

    foreach ($line in $ReqString){
        $lineCount++
        if($found){
            if(($line.length -gt 0) -and (IsNumeric($line.substring(4,1)))){
	            $obj="" | Select ExtensionOID,ExtensionName,ExtensionField
                $obj.ExtensionOID=($line.Split(":")[0]).trim()
                $obj.ExtensionName=$ReqString[$lineCount].trim()
                $Count++
                $fieldNum=0
                $i=1
                do{
                    $fieldLine=$ReqString[$lineCount+$i++]
                    $obj.ExtensionField+=$fieldLine.trim()+"`r`n"
                }until(($ReqString[$lineCount+$i].length -eq 0) -or !($ReqString[$lineCount+$i].StartsWith(" ")))
                $Extensions += $obj
            }
        }
        if($line -match "Certificate Extensions:"){
            [int32]$numCount=$line.Split(":")[1]
            if($numCount -gt 0){$found=$true}
        }
        if($found -and ($Count -eq $numCount)){$found=$false}
    }
    return $Extensions
}

function Get-SubjectName
{
    param(
        [Parameter(mandatory=$true)] [Array]$ReqString
    )
    #
    # return data type: custom PS object
    # object contain of: CN,OU,O,L,ST,STREET,C,E
    #

    $line=""
    $lineCount=0
    $found=$false
    $fieldNum=0
    $fieldLine=""
    $Extensions=@()
    $numCount=0
    $Count=0
    $i=0
 	$obj="" | Select CN,OU,O,L,S,STREET,C,E

    foreach ($line in $ReqString){
        if($found){
            do{
                $fieldLine=$ReqString[$lineCount++]
                $SubjectType=$fieldLine.Split("=")[0].trim()
                switch($SubjectType)
                {
                    "CN" {$obj.CN=$fieldLine.Split("=")[1].trim()}
                    "OU" {$obj.OU=$fieldLine.Split("=")[1].trim()}
                    "O" {$obj.O=$fieldLine.Split("=")[1].trim()}
                    "L" {$obj.L=$fieldLine.Split("=")[1].trim()}
                    "S" {$obj.S=$fieldLine.Split("=")[1].trim()}
                    "STREET" {$obj.STREET=$fieldLine.Split("=")[1].trim()}
                    "C" {$obj.C=$fieldLine.Split("=")[1].trim()}
                    "E" {$obj.E=$fieldLine.Split("=")[1].trim()}
                    Default {$found=$false}
                }
            }while($found)
        }
        $lineCount++
        if($line -match "Subject:"){$found=$true}
    }
    return $obj
}

function Get-ClientInfo
{
    param(
        [Parameter(mandatory=$true)] [Array]$ReqString
    )
    #
    # return data type: tbd
    #

}

function Get-CertTemplate
{
    param(
        [Parameter(mandatory=$true)]$arrExt,
        [Parameter(mandatory=$true)]$arrAttrib
    )
    #
    # return data type: array
    # array contain of: Template information from Certificate Extensions (if available) and from Request Attributes (if Available)
    # if the information are taken from Certificate Extensions, only the template OID is available; if taken from Request Attributes,
    # a template name is available.
    # if none of the above is available in the request file, an array with length of zero is returned.
    # 

    $arrCertTmpl=@()

    if ($arrExt) {
        Foreach($obj in $arrExt) {
            if($obj.ExtensionOID -match "1.3.6.1.4.1.311.21.7"){$arrCertTmpl+=$obj.ExtensionField.Split("`r`n")[0],"OID"}
        }
    }
    if ($arrAttrib) {
        Foreach($obj in $arrAttrib) {
            if($obj.AttribField1 -match "CertificateTemplate"){$arrCertTmpl+=$obj.AttribField1,"Name"}
        }
    }
    return $arrCertTmpl
}

function Get-CSP
{
    param(
        [Parameter(mandatory=$true)]$arrExt,
        [Parameter(mandatory=$true)]$arrAttrib
    )
    #
    # return data type: array
    # array contain of: max. five lines of CSP information
    #

    $CSP=@()
    if($arrAttrib.count -ne 0)
    {
        Foreach($obj in $arrAttrib) {
            if($obj.AttributeName -match "CSP"){
                if($obj.AttribField1 -match "Provider ="){$CSP+=$obj.AttribField1.split("=")[1]}
                if($obj.AttribField2 -match "Provider ="){$CSP+=$obj.AttribField2.split("=")[1]}
                if($obj.AttribField3 -match "Provider ="){$CSP+=$obj.AttribField3.split("=")[1]}
                if($obj.AttribField4 -match "Provider ="){$CSP+=$obj.AttribField4.split("=")[1]}
                if($obj.AttribField5 -match "Provider ="){$CSP+=$obj.AttribField5.split("=")[1]}
            }
        }
    }
    return $CSP
}

function Get-EnhancedKeyUsage
{
    param(
        [Parameter(mandatory=$true)]$arrExt
    )
    #
    # return data type: array
    # array contain of: all Enhanced Key Usages mentioned in the Certificate Extensions
    # I am not sure if Enhanced Key Usages can be defined as Request Attributes (never seen this)...
    #

    $EnhKeyUsage=@()
 	$Ekuobj="" | Select Name,OID
    if($arrExt.count -ne 0)
    {
        Foreach($obj in $arrExt) {
            if($obj.ExtensionOID -match "2.5.29.37"){
                $EkuList=$obj.ExtensionField.Split(")")
                $EkuList|ForEach-Object{
 	                $Ekuobj="" | Select Name,OID
                    if(!($_.trim() -eq "")){
                        $Ekuobj.Name=$_.split("(")[0]
                        $Ekuobj.OID=$_.split("(")[1]
                        $EnhKeyUsage+=$Ekuobj
                    }
                }
            }
        }
    }
    return $EnhKeyUsage
}

function Get-KeyUsage
{
    param(
        [Parameter(mandatory=$true)]$arrExt
    )
    #
    # return data type: string
    # string contain of: Key Usage taken from from the Certificate Extensions
    # I am not sure if  Key Usages can be defined as Request Attributes (never seen this)...
    #

    $KeyUsage=""
    if($arrExt.count -ne 0)
    {
        Foreach($obj in $arrExt) {
            if($obj.ExtensionOID -match "2.5.29.15"){
                $KeyUsage=$obj.ExtensionField
            }
        }
    }
    return $KeyUsage
}

function Get-SAN
{
    param(
        [Parameter(mandatory=$true)]$arrExt
    )
    #
    # return data type: array of custom PS objects
    # objects contain of: Type, SAN
    #

    $temp=""
    $SANTemp=""
    $SAN=@()
    $SANEntry="" | select Type, SAN
    if($arrExt.count -ne 0)
    {
        Foreach($obj in $arrExt) {
            if($obj.ExtensionOID -match "2.5.29.17"){
                $temp=$obj.ExtensionField.trim().split("`n")
                foreach($SANTemp in $temp){
                    if ($SANTemp -match "=") {
                        $SANEntry="" | select Type, SAN
                        $SANEntry.Type=$SANTemp.split("=")[0]
                        $SANEntry.SAN=$SANTemp.split("=")[1].trim()
                        $SAN+=$SANEntry
                    }
                }
            }
        }
    }
    return $SAN
}

function Get-SignAlgorithm
{
    param(
        [Parameter(mandatory=$true)] [Array]$ReqString
    )
    #
    # return data type: custom PS object
    # object contain of: AlgoOID,AlgoName
    #

    $line=""
    $found=$false
 	$obj="" | Select AlgoOID,AlgoName

    foreach ($line in $ReqString){
        if($found){
            if($line -match "    Algorithm ObjectId:"){
                $obj.AlgoOID=($line.Split(":")[1].trim()).Split(" ")[0]
                $obj.AlgoName=($line.Split(":")[1].trim()).Split(" ")[1]
            }
            $found=$false
        }
        if($line -match "Signature Algorithm:"){$found=$true}
    }
    return $obj
}

function Get-KeyAlgorithm
{
    param(
        [Parameter(mandatory=$true)] [Array]$ReqString
    )
    #
    # return data type: custom PS object
    # object contain of: AlgoOID,AlgoName
    #

    $line=""
    $found=$false
 	$obj="" | Select AlgoOID,AlgoName

    foreach ($line in $ReqString){
        if($found){
            if($line -match "    Algorithm ObjectId:"){
                $obj.AlgoOID=($line.Split(":")[1].trim()).Split(" ")[0]
                $obj.AlgoName=($line.Split(":")[1].trim()).Split(" ")[1]
            }
            $found=$false
        }
        if($line -match "Public Key Algorithm:"){$found=$true}
    }
    return $obj
}

function Get-KeyLength
{
    param(
        [Parameter(mandatory=$true)] [Array]$ReqString
    )
    #
    # return data type: string
    #

    $line=""
    $found=$false
 	$KeyLength=""

    foreach ($line in $ReqString){
        if($line -match "Public Key Length:"){$KeyLength=($line.Split(":")[1].trim()).split(" ")[0]}
    }
    return [int32]$KeyLength
}

Function Dump-Request
{
    param(
        [Parameter(mandatory=$true)] $ReqFileName
    )

    $CsrObj = "" | select SubjectName,ClientInfo,SAN,CertTmpl,KeyUsage,EnhancedKeyUsage,Extensions,RequestAttributes,KeyAlgo,KeyLength,SignAlgo,EnrollmentCSP,SignatureMatch

    $certReq = certutil $ReqFileName

    $CsrObj.SignatureMatch = Check-Signature($certReq)
    $CsrObj.RequestAttributes = Get-RequestAttributes($certReq)
    $CsrObj.Extensions = Get-CertExtensions($certReq)
    $CsrObj.subjectName = Get-SubjectName($certReq)
    $CsrObj.clientInfo = Get-ClientInfo($certReq)
    if ($CsrObj.Extensions) {
        $CsrObj.certTmpl = Get-CertTemplate $CsrObj.Extensions $CsrObj.RequestAttributes
    }
    $CsrObj.enrollmentCSP = Get-CSP $CsrObj.Extensions $CsrObj.RequestAttributes
    $CsrObj.enhancedKeyUsage = Get-EnhancedKeyUsage($CsrObj.Extensions)
    $CsrObj.keyUsage = Get-KeyUsage($CsrObj.Extensions)
    $CsrObj.SAN = Get-SAN($CsrObj.Extensions)
    $CsrObj.signAlgo = Get-SignAlgorithm($certReq)
    $CsrObj.KeyAlgo = Get-KeyAlgorithm($certReq)
    $CsrObj.KeyLength = Get-KeyLength($certReq)

    return $CsrObj
}



#sample checks for compliance...
#$passedDNSCheck=check-DNSSubjectNames $subjectName $SAN
#$subjectPassedRegExCheck=RegEx-SubjectCheck $subjectName
#$SANPassedRegExCheck=RegEx-SANCheck $SAN

#$ReqFile


