function Get-ProgisticsComputerName {
    $Script:ProgisticsComputerName
}

function Set-ProgisticsComputerName {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    process {
        $Script:ProgisticsComputerName = $ComputerName
        $Script:Proxy = New-WebServiceProxy -Uri "http://$ComputerName/amp/wsdl" -Class Progistics -Namespace Progistics
    }    
}

function Get-ProgisticsWebServiceProxy {
    $Script:Proxy
}

function Invoke-ProgisticsAPI {
    param (
        $MethodName,
        $Parameter
    )
    $Proxy = Get-ProgisticsWebServiceProxy
    $Result = $Proxy.$MethodName($Parameter)
    $Result.result.resultData
}

function Get-ProgisticsCarriers {
    Invoke-ProgisticsAPI -MethodName ListCarriers -Parameter (New-Object Progistics.ListCarriersRequest)
}

function Find-ProgisticsPackage {
    param (
        [Parameter(Mandatory)]$carrier,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$company,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$address1,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$city,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$postalCode,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$phone
    )
    begin {
        $PSBoundParameters.Remove("carrier") | Out-Null
    }
    process {
        $Request = New-Object Progistics.SearchRequest -Property @{
            carrier = $carrier
            filters = New-Object Progistics.DataDictionary -Property @{
                consignee = New-Object Progistics.NameAddress -Property (
                    $PSBoundParameters + @{
                        countryCode = "US"
                    }
                )
            }
        }

        $Result = Invoke-ProgisticsAPI -MethodName Search -Parameter $Request
        $Result.resultdata.consignee    
    }
}