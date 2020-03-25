$working_directory = "C:\work"
$latest_folder = Get-ChildItem -Directory -path $working_directory | Sort-Object $_.lastwritetime | Select-Object -First 1
$XML_location = $working_directory + "\" + $latest_folder.Name + "\XML Data"
$xml_file = $XML_location + "\All_systems_formatted.xml"

[xml]$inputfile = Get-Content $xml_file

#Create Directory for CSVs
if (!(Test-Path "$($PSScriptRoot)\CSV"))
{
    Write-Host "Creating Directory Structure for CSV"
    $NULL = New-Item -Path "$($PSScriptRoot)\CSV" -ItemType Directory
}

function Get-DataFields {
    Param ($file)

    $selection =@()

    switch ($file) {
        "clusterdetails" {
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'ClusterSerialNumber'
            $selection += 'OntapVersion'
            $selection += 'AvailableOntapVersion'
            $selection += 'NodeCount'
            $selection += 'DataVserverCount'
            $selection += 'ClusterHaConfigured'
            $selection += 'SwitchlessClusterEnabled'
            $selection += 'Timezone'
            $selection += 'ClusterRawCapacityinBytes'
            break
        }
        "clusterstoragesummary"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'SystemModel'
            $selection += 'TotalNodes'
            $selection += 'TotalShelves'
            $selection += 'TotalDrives'
            $selection += 'TotalVservers'
            $selection += 'TotalLifs'
            $selection += 'TotalAggregates'
            $selection += 'TotalAllocatedCapacityInBytes'
            $selection += 'TotalAllocatedUsedCapacityInBytes'
            $selection += 'TotalAllocatedAvailableCapacityInBytes'
            $selection += 'TotalVolumeCapacityInBytes'
            $selection += 'TotalVolumeUsedCapacityInBytes'
            $selection += 'TotalDedupeSavingsInBytes'
            $selection += 'TotalCompressionSavingsInBytes'
            $selection += 'TotalSavingsInBytes'
            break
        }
        "nodedetails"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'NodeName'
            $selection += 'SystemModel'
            $selection += 'SystemSerialNumber'
            $selection += 'SystemId'
            $selection += 'PartnerSystemName'
            $selection += 'PartnerSystemSerialNumber'
            $selection += 'PartnerSystemId'
            $selection += 'OntapVersion'
            $selection += 'AvailableOntapVersion'
            $selection += 'FirmwareRevision'
            $selection += 'IsEpsilonNode'
            $selection += 'IsAllFlashOptimized'
            $selection += 'NodeUptime'
            $selection += 'NodeLocation'
            break
        }
        "shelfdetails"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += @{n="NodeNames";e={$_.nodenames.split() -join ','}}
            $selection += 'SystemModel'
            $selection += 'ShelfName'
            $selection += 'ShelfId'
            $selection += 'ShelfUid'
            $selection += 'SerialNumber'
            $selection += 'ShelfState'
            $selection += 'ShelfModel'
            $selection += 'ShelfType'
            $selection += 'FirmwareRevision'
            $selection += 'ShelfBayCount'
            $selection += 'DrivesPerBay'
            $selection += 'DriveSlotCount'
            $selection += @{n="DriveOwnership";e={$_.DriveOwnership.split() -join ','}}
            $selection += 'DriveTypeAndCount'
            $selection += 'FailedDiskCount'
            $selection += 'MissingDiskCount'
            break
        }
        "drivedetails"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += @{n="NodeNames";e={$_.nodenames.split() -join ','}}
            $selection += 'DiskName'
            $selection += 'Shelf'
            $selection += 'ShelfBay'
            $selection += 'ShelfSerialNumber'
            $selection += 'Vendor'
            $selection += 'Model'
            $selection += 'DiskType'
            $selection += 'MarketingCapacity'
            $selection += 'DiskRpm'
            $selection += 'SerialNumber'
            $selection += 'FirmwareRevision'
            $selection += 'HomeNodeName'
            $selection += 'OwnerNodeName'
            $selection += 'RootOwner'
            $selection += 'DataOwner'
            $selection += @{n="DiskPathNames";e={$_.diskpathnames.split() -join ','}}
            $selection += 'PowerOnDuration'
            break
        }
        "aggregateconfiguration"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'AggregateName'
            $selection += 'HomeNodeName'
            $selection += 'OwnerNodeName'
            $selection += 'State'
            $selection += 'RaidStatus'
            $selection += 'RaidType'
            $selection += 'DiskCount'
            $selection += @{n="DiskCountByType";e={$_.DiskCountByType.split() -join ','}}
            $selection += 'RaidGroupSize'
            $selection += 'RaidGroupCount'
            $selection += 'PercentSnapshotSpace'
            $selection += 'SizeNominalInBytes'
            $selection += 'FlexVolCount'
            $selection += 'SnapshotSchedule'
            $selection += 'SnapshotCount'
            break
        }
        "aggregatespacedetails"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'AggregateName'
            $selection += 'PercentUsedCapacity'
            $selection += 'SizeNominalInBytes'
            $selection += 'SizeTotalInBytes'
            $selection += 'SizeUsedInBytes'
            $selection += 'SizeAvailableInBytes'
            $selection += 'PhysicalUsedPercent'
            $selection += 'PhysicalUsed'
            break
        }
        "svmconfiguration"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'VserverName'
            $selection += 'VserverType'
            $selection += 'AdminState'
            $selection += 'OperationalState'
            $selection += 'AllowedProtocols'
            $selection += 'NameServerSwitch'
            $selection += 'NameMappingSwitch'
            $selection += 'SnapshotPolicy'
            $selection += 'QuotaPolicy'
            $selection += 'AntivirusOnAccessPolicy'
            $selection += 'RootVolumeSecurityStyle'
            $selection += 'Language'
            $selection += 'RootVolume'
            $selection += 'RootVolumeAggregate'
            $selection += 'Ipspace'
            $selection += 'IsDomainAuthTunnel'
            break
        }
        "flexvolspacedetails"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'VserverName'
            $selection += 'VolumeType'
            $selection += 'State'
            $selection += 'SizeNominalInBytes'
            $selection += 'SizeTotalInBytes'
            $selection += 'SizeUsedInBytes'
            $selection += 'SizeAvailableInBytes'
            $selection += 'PercentageSizeUsed'
            $selection += 'SnapshotReserveSizeInBytes'
            $selection += 'SizeUsedBySnapshotsInBytes'
            $selection += 'TotalDedupeSavingsInBytes'
            $selection += 'PercentageDeduplicationSpaceSaved'
            $selection += 'TotalCompressionSavingsInBytes'
            $selection += 'PercentageCompressionSpaceSaved'
            $selection += 'TotalSavingsInBytes'
            $selection += 'PercentageTotalSpaceSaved'
            break
        }
        "flexvolconfiguration"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'VserverName'
            $selection += 'VolumeName'
            $selection += 'Aggregate'
            $selection += 'Type'
            $selection += 'State'
            $selection += 'UnixPermissions'
            $selection += 'LanguageCode'
            $selection += 'SnapshotPolicy'
            $selection += 'ExportPolicy'
            $selection += 'SecurityStyle'
            $selection += 'IsFilesysSizeFixed'
            $selection += 'IsCloneVol'
            $selection += 'HasLuns'
            $selection += 'DpSnapmirrorDestinationCount'
            $selection += 'VaultSnapmirrorDestinationCount'
            $selection += 'XdpSnapmirrorDestinationCount'
            break
        }
        "networkportsettings"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'NodeName'
            $selection += 'Port'
            $selection += 'LinkStatus'
            $selection += 'PortType'
            $selection += 'Role'
            $selection += 'Mtu'
            $selection += 'FlowControl'
            $selection += 'Speed'
            $selection += 'Ipspace'
            $selection += 'BroadcastDomain'
            break
        }
        "networkportinterfacegroups"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'NodeName'
            $selection += 'IfgrpName'
            $selection += 'Mode'
            $selection += 'DistributionFunction'
            $selection += 'MACAddress'
            $selection += 'PortParticipation'
            $selection += 'Ports'
            $selection += 'UpPorts'
            break
        }
        "networkportvlansettings"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'NodeName'
            $selection += 'InterfaceName'
            $selection += 'VlanID'
            $selection += 'ParentInterface'
            $selection += 'GVRPEnabled'
            break
        }
        "networklifsettings"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'VserverName'
            $selection += 'InterfaceName'
            $selection += 'Role'
            $selection += 'Status'
            $selection += 'DataProtocols'
            $selection += 'IPAddress'
            $selection += 'CurrentNode'
            $selection += 'CurrentPort'
            $selection += 'IsHome'
            $selection += 'HomeNode'
            $selection += 'HomePort'
            $selection += 'IsAutoRevert'
            $selection += 'RoutingGroupName'
            $selection += 'FirewallPolicy'
            $selection += 'FailoverPolicy'
            $selection += @{n="FailoverTarget";e={$_.failovertarget.split() -join ','}}
            break
        }
        "exportpolicyrules"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'VserverName'
            $selection += 'PolicyName'
            $selection += 'RuleIndex'
            $selection += 'ClientMatch'
            $selection += 'Protocol'
            $selection += 'RORule'
            $selection += 'RWRule'
            $selection += 'AnonUserId'
            $selection += 'SuperUser'
            break
        }
        "cifsservers"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'VserverName'
            $selection += 'CifsServer'
            $selection += 'Domain'
            $selection += 'DomainNetBIOSName'
            $selection += 'WinsServers'
            $selection += 'PreferredDC'
            $selection += 'GPOEnabled'
            $selection += 'HomeDirAccessforAdminEnabled'
            $selection += 'HomeDirAccessforPublicEnabled'
            break
        }
        "hardwarelifecycleinformation"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += 'DeviceType'
            $selection += 'Model'
            $selection += 'PartNumber'
            $selection += 'Count'
            $selection += 'EOA'
            $selection += 'EOSHW'
            break
        }


        default {"error $($file)"}
    }
    return $selection
}

$worksheet = @()
foreach ($worksheet in $(Get-Content .\files.txt).tolower()){
    $selection_result = Get-DataFields $worksheet
    if ($selection_result -eq "error"){
        write-host "Error"
    }else{
        $inputfile.'NetAppDocs.ONTAP.BuildDoc'.$worksheet | Select-Object -Property $selection_result | Export-Csv $(".\csv\$($worksheet).csv") -NoTypeInformation -Delimiter:","
    }


}