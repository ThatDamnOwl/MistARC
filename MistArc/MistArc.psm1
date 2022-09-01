## Imports

Import-module PowerArcGIS -force
Import-module PowerMist -force

## Basic variables and getters/setters

$MistARCServiceName = "MistService"
$MistARCEnabledSites = @()
$ModuleFolder = (Get-Module MistArc -ListAvailable).path -replace "MistArc\.psm1"

Function Set-MistARCServiceName
{
    param
    (
        [string]
        $NewServiceName
    )
    set-variable -scope 1 -name MistARCServiceName -value $NewServiceName
}

Function Set-MistARCEnabledSites
{
    param
    (
        $NewEnabledSites
    )
    set-variable -scope 1 -name MistARCEnabledSites -value $NewEnabledSites
}

Function Get-MistARCServiceName
{
    return $MistARCServiceName
}

Function Get-MistARCEnabledSites
{
    return $MistARCEnabledSites
}

## Basic functions


Function Invoke-MistARCVariableSave 
{
    $AllVariables = Get-Variable -scope 1 | where {$_.name -match "MistARC"}
    $VariableStore = @{}
    foreach ($Variable in $AllVariables)
    {
        if ($Variable.value.GetType().name -eq "PSCredential")
        {
            $VariableStore += @{
                                   "username" = $Variable.value.username
                                   "securepass" = ($Variable.value.Password | ConvertFrom-SecureString)
                               }
        }
        else {
            $VariableStore += @{$Variable.name = $Variable.Value}
        }
    }

    $VariableStore.GetEnumerator() | export-csv "$ModuleFolder\$($ENV:Username)-Variables.csv"
}

Function Invoke-MistARCVariableLoad
{
    $VariablePath = "$ModuleFolder\$($ENV:Username)-Variables.csv"
    if (test-path $VariablePath)
    {
        $VariableStore = import-csv $VariablePath

        foreach ($Variable in $VariableStore)
        {
            if ($Variable.name -match "(username|securepass)")
            {
                if ($Variable.name -eq "username")
                {
                    Write-Debug "Importing MistARCCredential"
                    $EncString = ($VariableStore | where {$_.name -eq "securepass"}).Value | ConvertTo-SecureString
                    $Credential = New-Object System.Management.Automation.PsCredential($Variable.Value, $EncString)
                    set-variable -scope 1 -name MistARCCredential -value $Credential
                }
            }
            else
            {
                Write-Debug "Importing $($Variable.name)"
                set-variable -scope 1 -name $Variable.Name -value $Variable.Value
            }
        }
    }

}


## Actual functions

Function New-MistWAPFeature
{
    param
    (
        $name,
        $status,
        $HardwareType,
        $ManagementIP,
        $SerialNumber,
        $HardwareInfo,
        $ManagementURL,
        $Lat,
        $Lng
    )

    $Attributes = @{
        "name" = $name
        "status" = $status
        "hardwareyype" = $HardwareType
        "managementip" = $ManagementIP
        "serialnumber" = $SerialNumber
        "hardwareinfo" = $HardwareInfo
        "managementurl" = $ManagementURL
    }

    $geometry = Convert-LatLngtoWebMerc $Lat $Lng

    return @{"attributes" = $Attributes; "geometry" = $geometry}
}

Function Sync-MistArc
{
    param
    (
        $Session
    )

    Write-Verbose "Getting all mist information"
    $Org = get-MistOrganizations $Session
    $Sites = Get-MistSites -Session $Session -OrgID $Org[0].org_id
    $AllMistDevices = Get-MistInventory $session $Org[0].org_id
    $ActiveMistDevices = Get-MistDeviceStats $session $Org[0].org_id
    $Offsets = import-csv "$MistArcModuleFolder\offsets-variables.csv"


    Write-Verbose "Getting all Mist Site Maps"
    
    $AllSiteMaps = @()
    
    foreach ($Site in $Sites)
    {
        $SiteMaps = Get-MistSiteMaps -Session $Session -SiteID $Site.id 
        foreach ($SiteMap in $SiteMaps)
        {
            $AllSiteMaps += $SiteMap
        }
    }

    Write-Verbose "Getting ArcGIS Feature entries"
    $ArcMistDevices = Get-FeatureServiceLayerFeatures -ServiceName $MistARCServiceName -LayerNumber 0 -all
    $SiteDevices = $ActiveMistDevices | where {$_.site_id -in $MistARCEnabledSites}

    foreach ($ArcDevice in $ArcMistDevices)# | where {$_.Attributes.SerialNumber -in $SiteDevices.Serial})
    {
        Write-Verbose "Checking device $($ArcDevice.attributes.name)"
        $MistDevice = $AllMistDevices | where {$_.Serial -eq $ArcDevice.Attributes.SerialNumber}
        $ActiveDevice = $ActiveMistDevices | where {$_.Serial -eq $ArcDevice.Attributes.SerialNumber}
        if ($ActiveDevice)
        {
            $Site = $Sites | where {$_.id -eq $ActiveDevice.site_id}
            $MercCoords = Convert-LatLngtoWebMerc -lat $Site.latlng.lat -lng $Site.latlng.lng

            $Offset = $Offsets | where {$_.site_id -eq $ActiveDevice.site_id -and $_.map_id -eq $ActiveDevice.map_id}

            if ($ActiveDevice.map_id.Length -gt 0 -and $Offset)
            {   
                Write-Verbose "Site offset found"
                Write-Verbose "Old X Coord = $($MercCoords.x)"
                Write-Verbose "Old Y Coord = $($MercCoords.y)"

                ## This offset stuff is kinda wonky might fix it eventually, I got this by placing three points on the
                ## rough area I planned to place WAPS (x,y) (x+10000,y) (x,y+10000) and measuring the actual distance 
                ## in meters between the points
                $MercCoords.x += $Offset.offset_x * (10000/8886)
                $MercCoords.y += $Offset.offset_y * (10000/8842)
                Write-Verbose "New X Coord = $($MercCoords.x)"
                Write-Verbose "New Y Coord = $($MercCoords.y)"

                $Map = $AllSiteMaps | where {$_.id -eq $ActiveDevice.map_id}

                $PPM = $Map.ppm
                if ($PPM -gt 0 -and $ActiveDevice.Serial)
                {
                    Write-Verbose "Map offset found"
                    Write-Verbose "The PPM for this map is $PPM"
                    Write-Verbose "The Device is located at $($ActiveDevice.x)"
                    Write-Verbose "The Device is located at $($ActiveDevice.y)"
                    Write-Verbose "Old X Coord = $($MercCoords.x)"
                    Write-Verbose "Old Y Coord = $($MercCoords.y)"
                    $WAPXOffset = ($ActiveDevice.x / $PPM) * (10000/8886)
                    $WAPYOffset = ($ActiveDevice.y / $PPM) * (10000/8842)

                    Write-Verbose "Map X Offset = $WAPXOffset"
                    Write-Verbose "Map Y Offset = $WAPYOffset"
                    if ($Offset)
                    {
                        $RadianAngle =  [system.Math]::PI * $Offset.Angle / 180 
                        if ($Offset.angle -lt 90)
                        {
                            Write-Verbose "Trying to combine Site and Map offset with angle $($Offset.angle)"
                            ## S=O/H C=A/H T=O/A
                            ## O=S*H A=C*H A=O/T O=A*T
                            $WAPXOffsetXComponent = [double]([double]$WAPXOffset * [System.math]::cos($RadianAngle))
                            Write-Verbose "XX Component = $WAPXOffsetXComponent"
                            $WAPXOffsetYComponent = [double]([double]$WAPXOffset * [System.math]::sin($RadianAngle))
                            Write-Verbose "XY Component = $WAPXOffsetYComponent"
                            $WAPYOffsetXComponent = [double]([double]$WAPYOffset * [System.math]::sin($RadianAngle)) * -1
                            Write-Verbose "YX Component = $WAPYOffsetXComponent"
                            $WAPYOffsetYComponent = [double]([double]$WAPYOffset * [System.math]::cos($RadianAngle))
                            Write-Verbose "YY Component = $WAPYOffsetYComponent"

                            $WAPYOffset = ($WAPXOffsetYComponent + $WAPYOffsetYComponent)
                            $WAPXOffset = ($WAPXOffsetXComponent + $WAPYOffsetXComponent)
                            Write-Verbose $WAPYOffset
                            Write-Verbose $WAPXOffset

                            $MercCoords.x += ($WAPXOffset)
                            $MercCoords.y += ($WAPYOffset * -1)
                            
                        }
                    }
                    else 
                    {    
                        $MercCoords.x += ($WAPYOffset * (10000/8886))
                        $MercCoords.y += (($WAPYOffset * -1) * (10000/8842))
                    }

                    Write-Verbose "New X Coord = $($MercCoords.x)"
                    Write-Verbose "New Y Coord = $($MercCoords.y)"
                }
            }
            $ArcDevice.geometry.x = $MercCoords.x
            $ArcDevice.geometry.y = $MercCoords.y
        }

        $ArcDevice.Attributes.name = $MistDevice.name
        $ArcDevice.Attributes.HardwareType = $MistDevice.type
        $ArcDevice.Attributes.HardwareInfo = $MistDevice.sku
        if ($MistDevice.site_id)
        {
            $ArcDevice.Attributes.ManagementURL = (Get-MistDeviceManagementURI $MistDevice)
        }

        if ($ActiveDevice)
        {
            $ArcDevice.Attributes.ManagementIP = $ActiveDevice.ip
            $ArcDevice.Attributes.status = 1

        }
        else {
            $ArcDevice.Attributes.status = 0
        }
    }
    #return $AllSiteMaps
    Invoke-FeatureServiceLayerFeatureUpdate -ServiceName $MistARCServiceName -LayerNumber 0 -UpdateJson (ConvertTo-Json $ArcMistDevices)
}


Function Import-MistActiveDevices
{
    param
    (
        $Session
    )
    $Org = get-MistOrganizations $Session
    $Sites = Get-MistSites -Session $Session -OrgID $Org[0].org_id
    $ActiveMistDevices = Get-MistDeviceStats $session $Org[0].org_id
    $MistInventory = Get-MistInventory $Session $Org[0].org_id
    $ArcMistDevices = Get-FeatureServiceLayerFeatures -ServiceName $MistARCServiceName -LayerNumber 0 -all
    $NewArcMistDevices = @()

    foreach ($ActiveMistDevice in $ActiveMistDevices)
    {
        $InventoryDevice = $MistInventory | where {$_.Serial -eq $ActiveMistDevice.Serial}
        #$ExistingArcMistDevice
        $ExistingArcMistDevice = $ArcMistDevices | where {$_.Attributes.SerialNumber -eq $ActiveMistDevice.Serial}
        if ($ExistingArcMistDevice.count -eq 0)
        {
            $Site = Get-MistSite -Session $Session -Siteid $ActiveMistDevice.site_id | where {$_.id -eq $ActiveMistDevice.site_id}

            $ArcMistDevice = New-MistWAPFeature -name $ActiveMistDevice.name `
                                                -status 1 `
                                                -HardwareType $InventoryDevice.type `
                                                -ManagementIP $ActiveMistDevice.ip `
                                                -SerialNumber $ActiveMistDevice.Serial `
                                                -HardwareInfo $InventoryDevice.sku `
                                                -ManagementURL (Get-MistDeviceManagementURI $ActiveMistDevice) `
                                                -Lat ($Site.latlng.lat) `
                                                -Lng ($Site.latlng.lng)
            $NewArcMistDevices += $ArcMistDevice
        }
        else {
            #$ExistingArcMistDevice.count
            #$ExistingArcMistDevice
            #read-host
        }
    }
    $NewArcMistDevices
}

Invoke-ARCGISVariableLoad