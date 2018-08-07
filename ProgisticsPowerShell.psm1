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
    $Response.result
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

function Get-ProgisticsService {
    param (
        $carrier,
        $shipper
    )
    Invoke-ProgisticsAPI -MethodName listServices -Property $PSBoundParameters
}

function Get-ProgisticsStock {
    Invoke-ProgisticsAPI -MethodName listStocks
}

function Get-ProgisticsDocumentFormat {
    param (
        $carrier
    )
    Invoke-ProgisticsAPI -MethodName listDocumentFormats -Property $PSBoundParameters
}

function Get-ProgisticsPrinterDevice {
    Invoke-ProgisticsAPI -MethodName listPrinterDevices
}

function Get-ProgisticsWindowsPrinter {
    Invoke-ProgisticsAPI -MethodName listWindowsPrinters
}

function Get-ProgisticsLocalPort {
    Invoke-ProgisticsAPI -MethodName listLocalPorts
}

function Remove-ProgisticsPackage {
    param (
        $Carrier,
        $MSN
    )
    $VoidPackagesRequest = New-Object Progistics.VoidPackagesRequest -Property @{
        carrier = $Carrier
        packages = [int[]]@($MSN)
    }
    Invoke-ProgisticsAPI -MethodName voidPackages -Parameter $VoidPackagesRequest
}

function Invoke-ProgisticsPackagePrint {
    param (
        $Carrier,
        $Shiper,
        $Document,
        [int]$MSN,
        $Output,
        $StockSymbol
    )
    "MSN" | 
    ForEach-Object {
        $PSBoundParameters.Remove($_) | Out-Null
    }

    $PrintRequest = New-Object Progistics.PrintRequest -Property (
        $PSBoundParameters + @{
            itemList = New-Object Progistics.PrintItemList -Property @{
                items = [System.Object[]]@($MSN)
                ItemsElementName = [Progistics.ItemsChoiceType[]]@([Progistics.ItemsChoiceType]::msn)
            }
            stock = New-Object Progistics.StockDescriptor -Property @{
                symbol = $StockSymbol
            }
        }
    )

    Invoke-ProgisticsAPI -MethodName Print -Parameter $PrintRequest
}

function New-ProgisticsPackage {
    param (
        [Parameter(ValueFromPipelineByPropertyName)]$Company,
        [Parameter(ValueFromPipelineByPropertyName,Mandatory)]$Address1,
        [Parameter(ValueFromPipelineByPropertyName)]$Address2,
        [Parameter(ValueFromPipelineByPropertyName,Mandatory)]$City,
        [Parameter(ValueFromPipelineByPropertyName,Mandatory)]$StateProvince,
        [Parameter(ValueFromPipelineByPropertyName,Mandatory)]$PostalCode,
        [Parameter(ValueFromPipelineByPropertyName)]$Residential,
        [Parameter(ValueFromPipelineByPropertyName)]$Phone,
        [Parameter(ValueFromPipelineByPropertyName)]$CountryCode,
        [Parameter(Mandatory)]$service,
        $Shipper,
        $Terms,
        $consigneeReference,
        [Parameter(Mandatory)]$WeightUnit,
        [Parameter(Mandatory)]$Weight,
        $ShipDate
    )
    $ConsigneeParameters = $PSBoundParameters | 
    ConvertFrom-PSBoundParameters -ExcludeProperty WeightUnit,Weight,consigneeReference,service,Shipper,Terms -AsHashTable

    $ShipRequest = New-Object Progistics.ShipRequest -Property @{
        service = $service
        defaults = New-Object Progistics.DataDictionary
        packages = [Progistics.DataDictionary[]]@(
            New-Object Progistics.DataDictionary -Property @{
                consignee = New-Object Progistics.NameAddress -Property (
                    $ConsigneeParameters
                )
                consigneeReference = $consigneeReference
                shipper = $Shipper
                terms = $Terms
                weight = New-Object Progistics.weight -Property @{
                    unit = $WeightUnit
                    amount = $Weight
                }
                shipdate = $ShipDate
            }
        )
    }
    
    Invoke-ProgisticsAPI -MethodName Ship -Parameter $ShipRequest
}