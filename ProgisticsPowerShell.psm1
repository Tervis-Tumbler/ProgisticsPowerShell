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
        $Parameter,
        $Property
    )
    $Proxy = Get-ProgisticsWebServiceProxy
    if (-not $Parameter) {
        if ($Property) {
            $Parameter = New-Object -TypeName Progistics."$($MethodName)Request" -Property $Property
        } else {
            $Parameter = New-Object -TypeName Progistics."$($MethodName)Request"
        }
    }
    $Response = $Proxy.$MethodName($Parameter)
    if ($Response.result.code -eq 0) {
        $Response.result.resultData
    } else {
        $Response.result
    }
}

function Get-ProgisticsCarrier {
    if (-not $Script:ProgisticsCarrier) {
        #https://connectship.com/docs/SDK/Technical_Reference/AMP_Reference/Core_Messages/Message_Elements/listCarriersRequest.htm
        $Script:ProgisticsCarrier = Invoke-ProgisticsAPI -MethodName ListCarriers
    }
    $Script:ProgisticsCarrier
}

function Find-ProgisticsPackage {
    param (
        [Parameter(Mandatory)]$carrier,
        [Parameter(Mandatory,ParameterSetName = "TrackingNumber")]$TrackingNumber,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName,ParameterSetName = "Consignee")]$company,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName,ParameterSetName = "Consignee")]$address1,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName,ParameterSetName = "Consignee")]$city,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName,ParameterSetName = "Consignee")]$postalCode,
        [Parameter(Mandatory,ValueFromPipelineByPropertyName,ParameterSetName = "Consignee")]$phone,
        [Switch]$globalSearch
    )
    begin {
        $PSBoundParameters.Remove("carrier") | Out-Null
        $PSBoundParameters.Remove("globalSearch") | Out-Null
    }
    process {
        #https://connectship.com/docs/SDK/Technical_Reference/AMP_Reference/Core_Messages/Message_Elements/searchRequest.htm
        #https://connectship.com/docs/SDK/Technical_Reference/AMP_Reference/Core_Messages/Complex_Types/DataDictionary.htm
        $Property = @{
            carrier = $carrier
            filters = New-Object Progistics.DataDictionary -Property $(
                if ($TrackingNumber) {
                    @{
                        trackingNumber = $TrackingNumber
                    }
                } else {
                    @{
                        consignee = New-Object Progistics.NameAddress -Property (
                            $PSBoundParameters + @{
                                countryCode = "US"
                            }
                        )
                    }
                }
            )
            globalSearch = $globalSearch
        }

        $Result = Invoke-ProgisticsAPI -MethodName Search -Property $Property
        $Result
    }
}

function Get-ProgisticsShipper {
    param (
        $carrier
    )
    Invoke-ProgisticsAPI -MethodName ListShippers -Property $PSBoundParameters
}

function Get-ProgisticsShipFile {
    param (
        $carrier,
        $shipper
    )
    Invoke-ProgisticsAPI -MethodName listShipFiles -Property $PSBoundParameters
}