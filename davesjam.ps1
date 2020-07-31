### Params ###
Param (
    [Parameter(Mandatory = $false, HelpMessage="Output File Path")][string] $csvOutputPath,                                             #Path ITFM Output File
    [Parameter(Mandatory = $false, HelpMessage="Reporting Period {YYYYMM Format}")][string] $reportingPeriod,                           #Billing Period
    [Parameter(Mandatory = $false, HelpMessage="Compute, Storage, and Network Discount %")][int32] $cpuMemStorageNetworkDiscount,       #Compute Discount %
    [Parameter(Mandatory = $false, HelpMessage="Interval Discount %")][int32] $intervalDiscount,                                        #Interval Discount % 
    [Parameter(Mandatory = $false, HelpMessage="License Discount %")][int32] $licenseDiscount,                                          #License Discount % 
    [Parameter(Mandatory = $false, HelpMessage="Extra Discount %")][int32] $extraDiscount                                               #Extra Discount %   
)

### DEBUG ###
$csvOutputPath = '/Users/rickynelson'
$reportingPeriod = '202007'
$cpuMemStorageNetworkDiscount = 2
$intervalDiscount = 2
$licenseDiscount = 2
$extraDiscount = 2

### Functions ###
function calculateQuantity {
    param (
        [Parameter(Mandatory = $true)][decimal] $cost,  
        [Parameter(Mandatory = $true)][decimal] $rate
    )

    return [decimal] ($cost / $rate)
}

function roundUnit {
    param (
        [Parameter(Mandatory = $true)][decimal] $unit  
    )

    return [decimal] ([Math]::Round($unit + 0.005,2))
}

function calculateDiscount {
    param (
        [Parameter(Mandatory = $true)][decimal] $cost,  
        [Parameter(Mandatory = $true)][decimal] $discount
    )

    return [decimal] ((roundUnit $cost) * ($discount * 0.01))
}

function calculateNet {
    param (
        [Parameter(Mandatory = $true)][decimal] $cost,  
        [Parameter(Mandatory = $true)][decimal] $discount
    )

    [decimal] $roundCost = (roundUnit $cost)
    return [decimal] ($roundCost - ($roundCost * ($discount * 0.01)))
}

function returnCurrency {
    param (
        [Parameter(Mandatory = $true)][decimal] $unit
    )

    return [string] $unit.ToString("C")
}

function buildItem {
    param (
        [Parameter(Mandatory = $true)][System.Object] $item
    )
    $itemArray = @()
    [array] $lineItemArray = @()

    [array] $tempLineItemArray

    #Process Storage
    $tempLineItemArray = @()
   
    #$tempLineItemArray += $item.lineItems | Select-Object -Property invoiceId, @{n='cloud';E={$item.cloud}}, @{n='resourceType';E={'Line Item'}}, @{n='resourceName';E={$_.refName}}, @{n='customerId';E={$item.customerId}}, @{n='projectId';E={$item.projectId}}, @{n='projectName';E={$item.projectName}}, @{n='plan';E={$_.usageType}}, @{n='consumptionCategory';E={$_.usageCategory}}, @{n='consumptionType';E={'N/A'}}, @{n='uom';E={$_.rateUnit}}, @{n='quantity';E={calculateQuantity $_.itemCost $_.itemRate}}, @{n='rate';E={$_.itemRate}}, @{n='cost';E={$_.itemCost}}, @{n='costNet';E={calculateNet $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='costDiscount';E={calculateDiscount $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='startDate';E={$_.startDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='endDate';E={$_.endDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='process';E={'No'}}  | Where-Object {$item.lineItems.invoiceId -Match $item.invoiceId}  #Build (VDI) Server Detail EndPoint Detail Info
    $itemArray += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'storage'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.storageCost}}, @{n='costNet';E={$_.storageCostNet}}, @{n='costDiscount';E={$_.storageCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}   #Build Server EndPoint Detail Info
    
    #Process Compute
    $itemArray += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'compute'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.computeCost}}, @{n='costNet';E={$_.computeCostNet}}, @{n='costDiscount';E={$_.computeCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build Server EndPoint Detail Info

    #Process Memory
    $itemArray += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'memory'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.computeCost}}, @{n='costNet';E={$_.memoryCostNet}}, @{n='costDiscount';E={$_.memoryCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build Server EndPoint Detail Info

    #Process Network
    $itemArray += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'network'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.networkCost}}, @{n='costNet';E={$_.networkCostNet}}, @{n='costDiscount';E={$_.networkCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build Server EndPoint Detail Info

    #Process License
    $itemArray += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'license'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.licenseCost}}, @{n='costNet';E={$_.licenseCostNet}}, @{n='costDiscount';E={$_.licenseCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build Server EndPoint Detail Info

    #Process Extra
    $itemArray += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'extra'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.extraCost}}, @{n='costNet';E={$_.extraCostNet}}, @{n='costDiscount';E={$_.extraCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build Server EndPoint Detail Info

    #Process Interval
    $itemArray += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'interval'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.intervalCost}}, @{n='costNet';E={$_.intervalCostNet}}, @{n='costDiscount';E={$_.intervalCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build Server EndPoint Detail Info

    
    return $itemArray
}

### Variables ###
[string] $morphUrl = "https://demo.morpheusdata.com/"                                               #Morpheus Instance
[string] $serviceBearer = ''                                    
[hashtable] $header = @{"Authorization" = "BEARER $serviceBearer"}

[string] $idPrefix = "UIS"                                                                          #Invoice ID Prefix                                                                              #Rounding Percision

[array] $endPoints = @()                                                                            #Will Store Clouds That We Want To Process
[array] $output = @()                                                                               #Variable To Hold All EndPoint Obtained From INVOICES API

$cloudAwsServer = @()                                                                               #Variable To Hold All AWS Server EndPoint Data Obtained From INVOICES API
[array] $cloudAwsVdi = @()                                                                          #Variable To Hold All AWS VDI EndPoint Data Obtained From INVOICES API
[array] $cloudAzureServer = @()                                                                     #Variable To Hold All Azure Server EndPoint Data Obtained From INVOICES API
[array] $cloudAzureVdi = @()                                                                        #Variable To Hold All Azure VDI EndPoint Data Obtained From INVOICES API
[array] $cloudOciServer = @()                                                                       #Variable To Hold All Oracle Cloud Server EndPoint Data Obtained From INVOICES API
[array] $cloudOciVdi = @()                                                                          #Variable To Hold All Oracle Cloud VDI EndPoint Data Obtained From INVOICES API

#Setting Up INVOICES API Call
$endPoints += 12                                                                                    #Example "zoneId" Show AWS: Cloud 3 (In Morpheus QA)/Cloud 12 (In Morpheus Demo); In API/CLI zoneId is Cloud
#$endPoints += 121                                                                                  #Example "zoneId" Show Azure: Cloud 22654 (In Morpheus QA)/Cloud 121 (In Morpheus Demo); In API/CLI zoneId is Cloud

#Iterate Through Each Defined Cloud And Build Invoice File Via INVOICES API and Capture OS Data Via SERVERS API
forEach ($endPoint In $endPoints) {
    try {
        #Run INVOICES API Query
        #Example "active" Filter For Active Objects 
        #Example "period" Show Reporting Period
        #Example "refType" Only Show Machine Type: ComputeServer (Virtual Machines)
        #Example "max" Export Number For Rows
        $invoicesEndPoints = "api/invoices?period=$($reportingPeriod)&max=100000&zoneId=" + $endPoint

        $responseStream = (Invoke-WebRequest -SkipCertificateCheck -Method Get -Uri ($morphUrl + $invoicesEndPoints) -Headers $header).content | ConvertFrom-Json | Select-Object -ExpandProperty invoices
        #Build Output Buffer
        $output += $responseStream | Select-Object `
        @{n='invoiceId';E={$_.id}}, `                                                               #Change id to invoiceId
        @{n='resourceName';E={$_.refName}}, `                                                       #Change refName to resourceName
        @{n='type';E={$_.refType}}, `                                                               #Change refType to type
        @{n='projectNumber';E={$_.project.tags.projectNumber}}, `                                   #projectNumber Of project.tag
        @{n='projectName';E={$_.project.name}}, `                                                   #projectName Of project
        @{n='projectId';E={$_.project.tags.projectId}}, `                                           #projectId >> projectId Of project.tag
        @{n='account';E={$_.project.tags.agencyName}}, `                                            #account >> agencyName Of project.tag
        @{n='cloud';E={$_.cloud.name}}, `                                                           #name Of cloud
        @{n='plan';E={$_.plan.name}}, `                                                             #name Of plan
        @{n='customerId';E={$_.project.tags.agency}}, `                                             #customerId >> agency Of project.tag
        @{n='startDate';E={$_.startDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, `               #startDate
        @{n='endDate';E={$_.endDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, `                   #endDate
        costType, `                                                                                 #costType
        @{n='storageCost';E={returnCurrency (roundUnit $_.storageCost)}}, `                                     #storageCost >> coneerted to local currency
        @{n='storageCostDiscount';E={returnCurrency (calculateDiscount $_.storageCost $cpuMemStorageNetworkDiscount)}}, ` #storageCostDiscount (calculated from storageCost and cpuMemStorageNetworkDiscount)
        @{n='storageCostNet';E={returnCurrency (calculateNet $_.storageCost $cpuMemStorageNetworkDiscount)}}, `      #storageCostNet (calculated from storageCost and storageCostDiscount)
        @{n='computeCost';E={returnCurrency (roundUnit $_.computeCost)}}, `                                     #computeCost >> coneerted to local currency
        @{n='computeCostDiscount';E={returnCurrency (calculateDiscount $_.computeCost $cpuMemStorageNetworkDiscount)}}, ` #computeCostDiscount (calculated from computeCost and cpuMemStorageNetworkDiscount)
        @{n='computeCostNet';E={returnCurrency (calculateNet $_.computeCost $cpuMemStorageNetworkDiscount)}}, `      #computeCostNet (calculated from computeCost and computeCostDiscount)
        @{n='memoryCost';E={returnCurrency roundUnit($_.memoryCost)}}, `                                       #memoryCost >> converted to local currency
        @{n='memoryCostDiscount';E={returnCurrency (calculateDiscount $_.memoryCost $cpuMemStorageNetworkDiscount)}}, ` #memoryCostDiscount (calculated from memoryCost and cpuMemStorageNetworkDiscount)
        @{n='memoryCostNet';E={returnCurrency (calculateNet $_.memoryCost $cpuMemStorageNetworkDiscount)}}, `        #memoryCostNet (calculated from memoryCost and memoryCostDiscount)
        @{n='networkCost';E={returnCurrency (roundUnit $_.networkCost)}}, `                                     #networkCost >> converted to local currency
        @{n='networkCostDiscount';E={returnCurrency (calculateDiscount $_.networkCost $cpuMemStorageNetworkDiscount)}}, ` #networkCostDiscount (calculated from networkCost and cpuMemStorageNetworkDiscount)
        @{n='networkCostNet';E={returnCurrency (calculateNet $_.networkCost $cpuMemStorageNetworkDiscount)}}, `      #networCostNet (calculated from networkCost and networkCostDiscount)
        @{n='licenseCost';E={returnCurrency (roundUnit $_.licenseCost)}}, `                                     #licenseCost >> converted to local currency
        @{n='licenseCostDiscount';E={returnCurrency (calculateDiscount $_.licenseCost $licenseDiscount)}}, `         #licenseCostDiscount (calculated from licenseCost and licenseDiscount)
        @{n='licenseCostNet';E={returnCurrency (calculateNet $_.licenseCost $licenseDiscount)}}, `                   #licenseCostNet (calculated from licenseCost and licenseCostDiscount)
        @{n='extraCost';E={returnCurrency (roundUnit $_.extraCost)}}, `                                         #extraCost >> converted to local currency
        @{n='extraCostDiscount';E={returnCurrency (calculateDiscount $_.extraCost $extraDiscount)}}, `               #extraCostDiscount (calculated from extraCost and extraDiscount)
        @{n='extraCostNet';E={returnCurrency (calculateNet $_.extraCost $extraDiscount)}}, `                         #extraCostNet (calculated from extraCost and extraCostDiscount)
        @{n='intervalCost';E={returnCurrency (roundUnit $_.intervalCost)}}, `                                   #intervalCost >> converted to local currency
        @{n='intervalCostDiscount';E={returnCurrency (calculateDiscount $_.intervalCost $intervalDiscount)}}, `      #intervalCostDiscount (calculated from intervalCost and intervalDiscount)
        @{n='intervalCostNet';E={returnCurrency (calculateNet $_.intervalCost $intervalDiscount)}}, `                #intervalCostNet (calculated from intervalCost and intervalCostDiscount)
        lineItems                                                                                   #lineItems {this contains the detail line items}
    }
    catch {
        Write-Output $_.Exception.Message
    }
}

### Iterate Through $output And Translate Cloud Name To Match KSE
forEach ($item in $output) {
    try {
        #Determine Object Type
        if ($item.type -match 'Instance' -or $item.type -eq 'ComputeServer' -or $item.type -eq 'Container') {
            $item.type ='Server'
        }
        elseif ($item.type -eq 'ComputeZone') {
            $item.type = 'Cloud'
        }
        else {
            $item.type = 'Service'
        }

        #Rename Clouds To Requested Format And Parse To Specific Cloud File
        if ($item.cloud -like "*aws*" -or $item.cloud -like "*amazon*") {
            $item.cloud = 'Amazon Web Services'                                                     #Rename cloud to Amazon Web Services
            
            #Seperate EndPoints Into Server And VDI Buckets Per Request
            if ($item.serverName -like "*win10*" -or $item.serverName -like "*vdi*") {              #VDI
                #VDI (Server/Host/Container)
                $cloudAwsVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'storage'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.storageCost}}, @{n='costNet';E={$_.storageCostNet}}, @{n='costDiscount';E={$_.storageCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAwsVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'compute'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.computeCost}}, @{n='costNet';E={$_.computeCostNet}}, @{n='costDiscount';E={$_.computeCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAwsVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'memory'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.computeCost}}, @{n='costNet';E={$_.memoryCostNet}}, @{n='costDiscount';E={$_.memoryCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAwsVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'network'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.networkCost}}, @{n='costNet';E={$_.networkCostNet}}, @{n='costDiscount';E={$_.networkCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAwsVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'license'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.licenseCost}}, @{n='costNet';E={$_.licenseCostNet}}, @{n='costDiscount';E={$_.licenseCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAwsVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'extra'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.extraCost}}, @{n='costNet';E={$_.extraCostNet}}, @{n='costDiscount';E={$_.extraCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAwsVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'interval'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.intervalCost}}, @{n='costNet';E={$_.intervalCostNet}}, @{n='costDiscount';E={$_.intervalCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                #Service
                $cloudAwsVdi += $item.lineItems | Select-Object -Property invoiceId, @{n='cloud';E={$item.cloud}}, @{n='resourceType';E={'Line Item'}}, @{n='resourceName';E={$_.refName}}, @{n='customerId';E={$item.customerId}}, @{n='projectId';E={$item.projectId}}, @{n='projectName';E={$item.projectName}}, @{n='plan';E={$_.usageType}}, @{n='consumptionCategory';E={$_.usageCategory}}, @{n='consumptionType';E={'N/A'}}, @{n='uom';E={$_.rateUnit}}, @{n='quantity';E={calculateQuantity $_.itemCost $_.itemRate}}, @{n='rate';E={$_.itemRate}}, @{n='cost';E={$_.itemCost}}, @{n='costNet';E={calculateNet $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='costDiscount';E={calculateDiscount $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='startDate';E={$_.startDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='endDate';E={$_.endDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='process';E={'No'}}  | Where-Object {$item.lineItems.invoiceId -Match $item.invoiceId}  #Build (VDI) Server Detail EndPoint Detail Info
            }
            else {                                                                                  #Server
                #Server/Host/Container
                


                    
                        # $dis = buildItem $item
                        # $($dis | where invoiceid -ne $null).count
                        # foreach($wat in $dis){
                        #     if($wat -ne $null){
                        #          $cloudAwsServer += $wat
                        #     }else{
                        #         Write-Output "null index 0"
                        #     }
                           
                        # }
                        $cloudAwsServer += $(buildItem $item | select -skip 0)
                    
                    

                                                                                      #Build Server EndPoint Detail Info
                #Service
                #$cloudAwsServer += $item.lineItems | Select-Object -Property invoiceId, @{n='cloud';E={$item.cloud}}, @{n='resourceType';E={'Line Item'}}, @{n='resourceName';E={$_.refName}}, @{n='customerId';E={$item.customerId}}, @{n='projectId';E={$item.projectId}}, @{n='projectName';E={$item.projectName}}, @{n='plan';E={$_.usageType}}, @{n='consumptionCategory';E={$_.usageCategory}}, @{n='consumptionType';E={'N/A'}}, @{n='uom';E={$_.rateUnit}}, @{n='quantity';E={calculateQuantity $_.itemCost $_.itemRate}}, @{n='rate';E={$_.itemRate}}, @{n='cost';E={$_.itemCost}}, @{n='costNet';E={calculateNet $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='costDiscount';E={calculateDiscount $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='startDate';E={$_.startDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='endDate';E={$_.endDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='process';E={'No'}}  | Where-Object {$item.lineItems.invoiceId -Match $item.invoiceId}  #Build Server Detail EndPoint Detail Info
            }
        }

        #Rename Clouds To Requested Format And Parse To Specific Cloud File
        if ($item.cloud -like "*azure*") {
            $item.cloud = 'Microsoft Azure'                                                         #Rename cloud to Microsoft Azure

            #Seperate EndPoints Into Server And VDI Buckets Per Request
            if ($item.serverName -like "*win10*" -or $item.serverName -like "*vdi*") {              #VDI
                #VDI (Server/Host/Container)
                $cloudAzureVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'storage'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.storageCost}}, @{n='costNet';E={$_.storageCostNet}}, @{n='costDiscount';E={$_.storageCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAzureVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'compute'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.computeCost}}, @{n='costNet';E={$_.computeCostNet}}, @{n='costDiscount';E={$_.computeCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAzureVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'memory'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.computeCost}}, @{n='costNet';E={$_.memoryCostNet}}, @{n='costDiscount';E={$_.memoryCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAzureVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'network'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.networkCost}}, @{n='costNet';E={$_.networkCostNet}}, @{n='costDiscount';E={$_.networkCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAzureVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'license'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.licenseCost}}, @{n='costNet';E={$_.licenseCostNet}}, @{n='costDiscount';E={$_.licenseCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAzureVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'extra'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.extraCost}}, @{n='costNet';E={$_.extraCostNet}}, @{n='costDiscount';E={$_.extraCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudAzureVdi += $item | Select-Object -Property invoiceId, cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'interval'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.intervalCost}}, @{n='costNet';E={$_.intervalCostNet}}, @{n='costDiscount';E={$_.intervalCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                #Service
                $cloudAzureVdi += $item.lineItems | Select-Object -Property invoiceId, @{n='cloud';E={$item.cloud}}, @{n='resourceType';E={'Line Item'}}, @{n='resourceName';E={$_.refName}}, @{n='customerId';E={$item.customerId}}, @{n='projectId';E={$item.projectId}}, @{n='projectName';E={$item.projectName}}, @{n='plan';E={$_.usageType}}, @{n='consumptionCategory';E={$_.usageCategory}}, @{n='consumptionType';E={'N/A'}}, @{n='uom';E={$_.rateUnit}}, @{n='quantity';E={calculateQuantity $_.itemCost $_.itemRate}}, @{n='rate';E={$_.itemRate}}, @{n='cost';E={$_.itemCost}}, @{n='costNet';E={calculateNet $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='costDiscount';E={calculateDiscount $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='startDate';E={$_.startDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='endDate';E={$_.endDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='process';E={'No'}}  | Where-Object {$item.lineItems.invoiceId -Match $item.invoiceId}  #Build (VDI) Server Detail EndPoint Detail Info
            }
            else {                                                                                  #Server
                #Server/Host/Container
                $cloudAzureServer += buildItem $item                                                #Build Server EndPoint Detail Info
                #Service
                #$cloudAzureServer += $item.lineItems | Select-Object -Property invoiceId, @{n='cloud';E={$item.cloud}}, @{n='resourceType';E={'Line Item'}}, @{n='resourceName';E={$_.refName}}, @{n='customerId';E={$item.customerId}}, @{n='projectId';E={$item.projectId}}, @{n='projectName';E={$item.projectName}}, @{n='plan';E={$_.usageType}}, @{n='consumptionCategory';E={$_.usageCategory}}, @{n='consumptionType';E={'N/A'}}, @{n='uom';E={$_.rateUnit}}, @{n='quantity';E={calculateQuantity $_.itemCost $_.itemRate}}, @{n='rate';E={$_.itemRate}}, @{n='cost';E={$_.itemCost}}, @{n='costNet';E={calculateNet $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='costDiscount';E={calculateDiscount $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='startDate';E={$_.startDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='endDate';E={$_.endDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='process';E={'No'}}  | Where-Object {$item.lineItems.invoiceId -Match $item.invoiceId}  #Build Server Detail EndPoint Detail Info
            }
        }

        #Rename Clouds To Requested Format And Parse To Specific Cloud File
        if ($item.cloud -like "*oracle*") {
            $item.cloud = 'Oracle Cloud'                                                            #Rename cloud to Oracle Cloud

            #Seperate EndPoints Into Server And VDI Buckets Per Request
            if ($item.serverName -like "*win10*" -or $item.serverName -like "*vdi*") {              #VDI
                #VDI (Server/Host/Container)
                $cloudOciVdi += $item | Select-Object -Property invoiceId, $item.cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'storage'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.storageCost}}, @{n='costNet';E={$_.storageCostNet}}, @{n='costDiscount';E={$_.storageCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudOciVdi += $item | Select-Object -Property invoiceId, $item.cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'compute'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.computeCost}}, @{n='costNet';E={$_.computeCostNet}}, @{n='costDiscount';E={$_.computeCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudOciVdi += $item | Select-Object -Property invoiceId, $item.cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'memory'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.computeCost}}, @{n='costNet';E={$_.memoryCostNet}}, @{n='costDiscount';E={$_.memoryCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudOciVdi += $item | Select-Object -Property invoiceId, $item.cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'network'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.networkCost}}, @{n='costNet';E={$_.networkCostNet}}, @{n='costDiscount';E={$_.networkCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudOciVdi += $item | Select-Object -Property invoiceId, $item.cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'license'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.licenseCost}}, @{n='costNet';E={$_.licenseCostNet}}, @{n='costDiscount';E={$_.licenseCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudOciVdi += $item | Select-Object -Property invoiceId, $item.cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'extra'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.extraCost}}, @{n='costNet';E={$_.extraCostNet}}, @{n='costDiscount';E={$_.extraCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                $cloudOciVdi += $item | Select-Object -Property invoiceId, $item.cloud, @{n='resourceType';E={$_.type}}, resourceName, customerId, projectId, projectName, plan, @{n='consumptionCategory';E={'interval'}}, @{n='consumptionType';E={$_.costType}}, @{n='uom';E={'N/A'}}, @{n='quantity';E={'N/A'}}, @{n='rate';E={'N/A'}}, @{n='cost';E={$_.intervalCost}}, @{n='costNet';E={$_.intervalCostNet}}, @{n='costDiscount';E={$_.intervalCostDiscount}}, startDate, endDate, @{n='process';E={'Yes'}}    #Build (VDI) Server EndPoint Detail Info
                #Service
                $cloudOciVdi += $item.lineItems | Select-Object -Property invoiceId, $item.cloud, @{n='resourceType';E={'Line Item'}}, @{n='resourceName';E={$_.refName}}, @{n='customerId';E={$item.customerId}}, @{n='projectId';E={$item.projectId}}, @{n='projectName';E={$item.projectName}}, @{n='plan';E={$_.usageType}}, @{n='consumptionCategory';E={$_.usageCategory}}, @{n='consumptionType';E={'N/A'}}, @{n='uom';E={$_.rateUnit}}, @{n='quantity';E={calculateQuantity $_.itemCost $_.itemRate}}, @{n='rate';E={$_.itemRate}}, @{n='cost';E={$_.itemCost}}, @{n='costNet';E={calculateNet $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='costDiscount';E={calculateDiscount $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='startDate';E={$_.startDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='endDate';E={$_.endDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='process';E={'No'}}  | Where-Object {$item.lineItems.invoiceId -Match $item.invoiceId}  #Build (VDI) Server Detail EndPoint Detail Info
            }
            else {                                                                                  #Server
                #Server/Host/Container
                $cloudOciServer += buildItem $item                                                  #Build Server EndPoint Detail Info
                #Service
                $cloudOciServer += $item.lineItems | Select-Object -Property invoiceId, $item.cloud, @{n='resourceType';E={'Line Item'}}, @{n='resourceName';E={$_.refName}}, @{n='customerId';E={$item.customerId}}, @{n='projectId';E={$item.projectId}}, @{n='projectName';E={$item.projectName}}, @{n='plan';E={$_.usageType}}, @{n='consumptionCategory';E={$_.usageCategory}}, @{n='consumptionType';E={'N/A'}}, @{n='uom';E={$_.rateUnit}}, @{n='quantity';E={calculateQuantity $_.itemCost $_.itemRate}}, @{n='rate';E={$_.itemRate}}, @{n='cost';E={$_.itemCost}}, @{n='costNet';E={calculateNet $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='costDiscount';E={calculateDiscount $_.itemCost $cpuMemStorageNetworkDiscount}}, @{n='startDate';E={$_.startDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='endDate';E={$_.endDate.ToUniversalTime().ToString('yyyy-MM-dd')}}, @{n='process';E={'No'}}  | Where-Object {$item.lineItems.invoiceId -Match $item.invoiceId}  #Build Server Detail EndPoint Detail Info
            }
        }
    }
    catch {
        Write-Output $_.Exception.Message
    }
}

### Export To Excel
try {
    [string] $outputFileAwsServer = $csvOutputPath + "/UNIS_CLOUD_CONSUMPTION_AWS_" + $reportingPeriod + "01-Server.csv"              #Autogenerate Filename Based On Desired Format
    $cloudAwsServer | Export-Csv $outputFileAwsServer                                                                                #Export AWS EndPoint Inventory To Excel
    [string] $outputFileAwsVdi = $csvOutputPath + "/UNIS_CLOUD_CONSUMPTION_AWS_" + $reportingPeriod + "01-Vdi.csv"                    #Autogenerate Filename Based On Desired Format
    $cloudAwsVdi | Export-Csv $outputFileAwsVdi                                                                                       #Export AWS EndPoint Inventory To Excel

    [string] $outputFileAzureServer = $csvOutputPath + "/UNIS_CLOUD_CONSUMPTION_AZURE_" + $reportingPeriod + "01-Server.csv"          #Autogenerate Filename Based On Desired Format
    $cloudAzureServer | Export-Csv $outputFileAzureServer                                                                             #Export Azure EndPoint Inventory To Excel
    [string] $outputFileAzureVdi = $csvOutputPath + "/UNIS_CLOUD_CONSUMPTION_AZURE_" + $reportingPeriod + "01-Vdi.csv"                #Autogenerate Filename Based On Desired Format
    $cloudAzureVdi | Export-Csv $outputFileAzureVdi                                                                                   #Export Azure EndPoint Inventory To Excel
    
    [string] $outputFileOciServer = $csvOutputPath + "/UNIS_CLOUD_CONSUMPTION_OCI_" + $reportingPeriod + "01-Server.csv"              #Autogenerate Filename Based On Desired Format
    $cloudOciServer | Export-Csv $outputFileOciServer                                                                                 #Export Azure EndPoint Inventory To Excel
    [string] $outputFileOciVdi = $csvOutputPath + "/UNIS_CLOUD_CONSUMPTION_OCI_" + $reportingPeriod + "01-Vdi.csv"                    #Autogenerate Filename Based On Desired Format
    $cloudOciVdi | Export-Csv $outputFileOciVdi                                                                                       #Export Azure EndPoint Inventory To Excel
}
catch {
    Write-Output $_.Exception.Message
}