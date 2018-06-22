function Get-ProgisticsComputerName {
    $Script:ProgisticsComputerName
}

function Set-ProgisticsComputerName {
    param (
        [Parameter(Mandatory,ValueFromPipelineByPropertyName)]$ComputerName
    )
    process {
        $Script:ProgisticsComputerName = $ComputerName
    }    
}

function Get-ProgisticsWebServiceProxy {
    if ($Script:Proxy) {
        $Script:Proxy
    } else {
        $ProgisticsComputerName = Get-ProgisticsComputerName
        $Proxy = New-WebServiceProxy -Uri "http://$ProgisticsComputerName/amp/wsdl" -Class Progistics -Namespace Progistics
    }
}

function Invoke-ProgisticsAPI {
    param (
        $MethodName,
        $Parameter
    )
    $Result = $Proxy.$MethodName($Parameter)
    $Result.result.resultData
}

function Get-ProgisticsCarriers {
    Invoke-ProgisticsAPI -MethodName ListCarriers -Parameter (New-Object Progistics.ListCarriersRequest)
}