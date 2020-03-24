$working_directory = "C:\work"
$latest_folder = Get-ChildItem -Directory -path $working_directory | Sort-Object $_.lastwritetime | select -First 1
$XML_location = $working_directory + "\" + $latest_folder.Name + "\XML Data"
$xml_file = $XML_location + "\All_systems_formatted.xml"

[xml]$inputfile = Get-Content $xml_file

function Get-Selection {
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
            $selection += 'NodeUptim'
            $selection += 'NodeLocation'
            break
        }
        "shelfdetails"{
            $selection += 'ClusterName'
            $selection += 'Location'
            $selection += 'Network'
            $selection += 'Environment'
            $selection += $("@{n=""NodeNames"";e={$_.nodenames.split() -join ','}}")
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
            $selection += $("@{n=""DriveOwnership"";e={$_.DriveOwnership.split() -join ','}}")
            $selection += 'DriveTypeAndCount'
            $selection += 'FailedDiskCount'
            $selection += 'MissingDiskCount'
            break
        }
        "drivedetails"{

            break
        }
        "aggregateconfiguration"{

            break
        }
        "aggregatespacedetails"{
        
            break
        }
        "svmconfiguration"{

            break
        }
        "aggregatestorageefficiency"{

            break
        }
        "flexvolspacedetails"{

            break
        }
        "flexvolconfiguration"{

            break
        }
        "networkportinterfacegroups"{

            break
        }
        "exportpolicyrules"{

            break
        }
        "cifsservers"{

            break
        }
        "portconfiguration"{

            break
        }
        "hardwarelifecycleinformation"{

            break
        }


        default {"error"}
    }
    return $selection
}

$worksheet = @()
foreach ($worksheet in $(Get-Content .\files.txt).tolower()){
    $selection_result = Get-Selection $worksheet
    if ($selection_result -eq "error"){
        write-host "Error"
    }else{
        $inputfile.'NetAppDocs.ONTAP.BuildDoc'.$worksheet | Select-Object -Property $selection_result | Export-Csv $(".\csv\$($worksheet).csv") -NoTypeInformation -Delimiter:","
    }


}