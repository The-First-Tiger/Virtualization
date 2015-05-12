namespace Virtualization.Extensions
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Management;
    using HyperV;

    internal static class ManagementExtensions
    {
        internal static ManagementScope GetScope(string hostname, ConnectionOptions connectionOptions = null)
        {
            if (connectionOptions != null)
            {
                return new ManagementScope(
                    new ManagementPath
                    {
                        Server = hostname,
                        NamespacePath = @"root\virtualization\v2"
                    },
                    connectionOptions);
            }

            return new ManagementScope(
                new ManagementPath
                {
                    Server = hostname,
                    NamespacePath = @"root\virtualization\v2"
                });
        }

        // TODO: Split into Add and Connect methods
        internal static void AddNetworAdapter(this ManagementScope scope, string vmIdentity, string switchName)
        {
            var virtualMachine = scope.GetVirtualMachine(vmIdentity);
            var virtualSystemSettingData = virtualMachine.GetVirtualSystemSettingData();

            var defaultRessourcePoolForSyntheticEthernetPorts = scope.GetDefaultRessourcePoolForSyntheticEthernetPorts();
            var allocationSettingData = defaultRessourcePoolForSyntheticEthernetPorts.GetAllocationSettingDataFromDefaultRessourcePool();

            allocationSettingData["VirtualSystemIdentifiers"] = new [] { Guid.NewGuid().ToString("B") };
            allocationSettingData["ElementName"] = "Network Adapter";
            allocationSettingData["StaticMacAddress"] = false;
            allocationSettingData.Put();

            var outParameters = scope.AddResourceSettings(virtualSystemSettingData, allocationSettingData);

            var resultingRessourceSettings = outParameters["ResultingResourceSettings"] as string[];

            var syntheticEthernetPortSettingsData = new ManagementObject(resultingRessourceSettings[0]);

            // --------------- Connect Adapter and Switch

            var defaultRessourcePoolForEthernetConnection = scope.GetDefaultRessourcePoolForEthernetConnection();
            var ethernetPortAllocationSettingData = defaultRessourcePoolForEthernetConnection.GetAllocationSettingDataFromDefaultRessourcePool();

            var virtualSwitch = scope.GetVirtualSwitch(switchName);

            ethernetPortAllocationSettingData["Parent"] = syntheticEthernetPortSettingsData;
            ethernetPortAllocationSettingData["HostResource"] = new [] { virtualSwitch };

            scope.AddResourceSettings(virtualSystemSettingData, ethernetPortAllocationSettingData);
        }

        internal static void StartVirtualMachine(this ManagementScope scope, string vmIdentity)
        {
            var virtualMachine = scope.GetVirtualMachine(vmIdentity);

            var inParameters = virtualMachine.GetMethodParameters("RequestStateChange");

            inParameters["RequestedState"] = (UInt16)VirtualMachineStates.Enabled;
            
            var outParameters = virtualMachine.InvokeMethod("RequestStateChange", inParameters, null);

            outParameters.HandleResult();            
        }

        internal static void HandleResult(this ManagementBaseObject invokedMethodResult)
        {
            if (!invokedMethodResult.WasSuccessful() && !invokedMethodResult.IsJob())
            {
                throw new InvalidOperationException("Could not execute!");
            }

            if (!invokedMethodResult.IsJob())
            {
                return;
            }

            var job = invokedMethodResult.GetCompletedJob();

            if (!job.WasJobSuccessful())
            {
                throw new InvalidOperationException("Job failed: " + job.GetJobErrorDescription());
            }
        }

        internal static string GetJobErrorDescription(this ManagementObject job)
        {
            return job["ErrorDescription"].ToString();
        }

        internal static bool WasJobSuccessful(this ManagementObject job)
        {
            return job.GetJobState() == JobStates.Completed;
        }

        internal static bool WasSuccessful(this ManagementBaseObject invokedMethodResult)
        {
            return ((UInt32)invokedMethodResult["ReturnValue"]) == 0;
        }

        internal static ManagementObject GetCompletedJob(this ManagementBaseObject invokedMethodResult)
        {
            var job = invokedMethodResult.GetJob();

            while (job.GetJobState() == JobStates.Starting
                   || job.GetJobState() == JobStates.Running)
            {
                job = invokedMethodResult.GetJob();
            }

            return job;
        }

        internal static ManagementObject GetJob(this ManagementBaseObject invokedMethodResult)
        {
            return new ManagementObject(invokedMethodResult["Job"].ToString());
        }

        internal static JobStates GetJobState(this ManagementObject job)
        {
            return (JobStates)(UInt16)job["JobState"];
        }

        internal static bool IsJob(this ManagementBaseObject invokedMethodResult)
        {
            return (UInt32)invokedMethodResult["ReturnValue"] == 4096;
        }

        internal static void SetVhd(this ManagementScope scope, string vmIdentity, string vhdFilepath)
        {
            var virtualMachine = scope.GetVirtualMachine(vmIdentity);
            var virtualSystemSettingData = virtualMachine.GetVirtualSystemSettingData();
            
            var diskDriveResource = scope.GetDiscDriveResource(virtualSystemSettingData);

            var allocationSettingData = scope.GetAllocationSettingDataForVirtualHardDisks();

            allocationSettingData["Parent"] = diskDriveResource;
            allocationSettingData["HostResource"] = new string[] { vhdFilepath };

            scope.AddResourceSettings(virtualSystemSettingData, allocationSettingData);            
        }

        internal static ManagementObject GetDiscDriveResource(this ManagementScope scope, ManagementObject virtualSystemSettingData)
        {
            var storageAllocationSettingData = scope.GetAllocationSettingDataForSyntheticDisDrives();

            var ideController = virtualSystemSettingData.GetIdeController();

            storageAllocationSettingData["Parent"] = ideController;
            storageAllocationSettingData["AddressOnParent"] = 0;

            var outParameters = scope.AddResourceSettings(virtualSystemSettingData, storageAllocationSettingData);

            var resultingRessourceSettings = outParameters["ResultingResourceSettings"] as string[];

            return new ManagementObject(resultingRessourceSettings[0]);
        }

        private static ManagementObject GetAllocationSettingDataFromDefaultRessourcePool(this ManagementObject defaultRessourcePool)
        {
            var allocationCapabilities = defaultRessourcePool.GetAllocationCapabilities();
            var defaultRelationshipClass = allocationCapabilities.GetDefaultRelationshipClass();

            return defaultRelationshipClass.GetAllocationSettingData();
        }

        internal static ManagementObject GetAllocationSettingDataForVirtualHardDisks(this ManagementScope scope)
        {
            var defaultRessourcePoolForVirtualHardDisks = scope.GetDefaultRessourcePoolForVirtualHardDisks();

            return defaultRessourcePoolForVirtualHardDisks.GetAllocationSettingDataFromDefaultRessourcePool();            
        }

        internal static ManagementObject GetAllocationSettingDataForSyntheticDisDrives(this ManagementScope scope)
        {
            var defaultRessourcePoolForSyntheticDiskDrives = scope.GetDefaultRessourcePoolForSyntheticDiskDrives();

            return defaultRessourcePoolForSyntheticDiskDrives.GetAllocationSettingDataFromDefaultRessourcePool();            
        }

        internal static void SetDynamicRam(this ManagementScope scope, string vmIdentity, int ramSize = 4096, int? reservation = 1024, int? limit = 8192)
        {
            var virtualMachine = scope.GetVirtualMachine(vmIdentity);

            var virtualSystemSettingData = virtualMachine.GetVirtualSystemSettingData();

            virtualSystemSettingData["VirtualNumaEnabled"] = false;
            virtualSystemSettingData.Put();

            scope.ModifySystemSettings(virtualSystemSettingData);

            var memorySettingData = virtualSystemSettingData.GetMemorySettingData();

            if (!reservation.HasValue)
            {
                reservation = ramSize / 4;
            }

            if (!limit.HasValue)
            {
                limit = ramSize * 2;
            }

            memorySettingData["DynamicMemoryEnabled"] = true;            
            memorySettingData["Reservation"] = reservation.Value;
            memorySettingData["VirtualQuantity"] = ramSize;
            memorySettingData["Limit"] = limit.Value;
            memorySettingData.Put();

            scope.ModifyResourceSettings(memorySettingData);
        }

        internal static void ModifySystemSettings(this ManagementScope scope, ManagementObject virtualSystemSettingData)
        {
            var virtualSystemManagementService = scope.GetVirtualSystemManagementService();

            var inParameters = virtualSystemManagementService.GetMethodParameters("ModifySystemSettings");
            var settingsText = virtualSystemSettingData.GetText(TextFormat.WmiDtd20);
            inParameters["SystemSettings"] = settingsText;

            virtualSystemManagementService.InvokeMethod("ModifySystemSettings", inParameters, null);
        }

        internal static ManagementBaseObject AddResourceSettings(
            this ManagementScope scope,
            ManagementObject virtualSystemSettingData,
            ManagementObject resourceSettingData)
        {
            var virtualSystemManagementService = scope.GetVirtualSystemManagementService();

            var inParameters = virtualSystemManagementService.GetMethodParameters("AddResourceSettings");

            inParameters["AffectedConfiguration"] = virtualSystemSettingData;
            inParameters["ResourceSettings"] = new string[]
                                               {
                                                   resourceSettingData.GetText(TextFormat.WmiDtd20)
                                               };

            return virtualSystemManagementService.InvokeMethod("AddResourceSettings", inParameters, null);
        }

        internal static void ModifyResourceSettings(this ManagementScope scope, ManagementObject resourceSettingsData)
        {
            var virtualSystemManagementService = scope.GetVirtualSystemManagementService();

            var inParameters = virtualSystemManagementService.GetMethodParameters("ModifyResourceSettings");
            var settingsText = resourceSettingsData.GetText(TextFormat.WmiDtd20);
            inParameters["ResourceSettings"] = new string[] { settingsText };

            virtualSystemManagementService.InvokeMethod("ModifyResourceSettings", inParameters, null);
        }

        internal static ManagementBaseObject DefineSystem(this ManagementScope scope)
        {
            var virtualSystemManagementService = scope.GetVirtualSystemManagementService();
            var definedSystem = virtualSystemManagementService.InvokeMethod("DefineSystem", virtualSystemManagementService.GetMethodParameters("DefineSystem"), null);

            if (definedSystem == null)
            {
                throw new InvalidOperationException("DefineSystem failed");
            }

            var returnValue = (uint)definedSystem["returnvalue"];

            if (returnValue != 0)
            {
                throw new InvalidOperationException("DefineSystem failed");
            }

            var vmPath = definedSystem["ResultingSystem"] as string;

            return new ManagementObject(vmPath);            
        }

        internal static void SetVirtualMachineName(this ManagementScope scope, string name, string vmIdentity)
        {
            var settings = scope.GetVirtualSystemSettingData(vmIdentity);

            settings["ElementName"] = name;
            settings.Put();

            scope.ModifySystemSettings(settings);
        }

        internal static string CreateVirtualMachine(this ManagementScope scope, string name)
        {
            var system = scope.DefineSystem();

            var vmIdentity = (string)system["name"];

            scope.SetVirtualMachineName(name, vmIdentity);            

            return vmIdentity;
        }

        internal static ManagementObject GetIdeController(this ManagementObject virtualSystemSettingData)
        {
            return virtualSystemSettingData.GetRelated("Msvm_ResourceAllocationSettingData")
                .OfType<ManagementObject>()
                .FirstOrDefault(
                    x =>
                       (UInt16)x["ResourceType"] == 5
                        && x["ResourceSubType"].ToString().Equals("Microsoft:Hyper-V:Emulated IDE Controller")
                        && x["Address"].ToString().Equals("0"));        
        }

        internal static ManagementObject GetAllocationSettingData(
            this ManagementObject defaultRelationshipClass)
        {
            return new ManagementObject(defaultRelationshipClass["PartComponent"].ToString()).Clone() as ManagementObject;
        }

        internal static ManagementObject GetDefaultRelationshipClass(this ManagementObject allocationCapabilities)
        {
            return allocationCapabilities.GetRelationships("Msvm_SettingsDefineCapabilities")
                .OfType<ManagementObject>()
                .FirstOrDefault(x => (UInt16)x["ValueRole"] == 0);
        }

        internal static ManagementObject GetAllocationCapabilities(
            this ManagementObject defaultRessourcePoolForSyntheticDiskDrives)
        {
            return defaultRessourcePoolForSyntheticDiskDrives.GetRelated(
                "Msvm_AllocationCapabilities",
                "Msvm_ElementCapabilities",
                null,
                null,
                null,
                null,
                false,
                null)
                    .OfType<ManagementObject>()
                    .FirstOrDefault();
        }

        private static ManagementObject GetDefaultRessourcePool(this ManagementScope scope, string resourceSubType)
        {
            return scope.GetRessourcePools()
                .FirstOrDefault(
                    x =>
                        x["ResourceSubType"].ToString().Equals(resourceSubType)
                        && (bool)x["Primordial"]);
        }

        internal static ManagementObject GetDefaultRessourcePoolForEthernetConnection(this ManagementScope scope)
        {
            return scope.GetDefaultRessourcePool("Microsoft:Hyper-V:Ethernet Connection");
        }

        internal static ManagementObject GetDefaultRessourcePoolForSyntheticEthernetPorts(this ManagementScope scope)
        {
            return scope.GetDefaultRessourcePool("Microsoft:Hyper-V:Synthetic Ethernet Port");           
        }

        internal static ManagementObject GetDefaultRessourcePoolForVirtualHardDisks(this ManagementScope scope)
        {
            return scope.GetDefaultRessourcePool("Microsoft:Hyper-V:Virtual Hard Disk");
        }

        internal static ManagementObject GetDefaultRessourcePoolForSyntheticDiskDrives(this ManagementScope scope)
        {
            return scope.GetDefaultRessourcePool("Microsoft:Hyper-V:Synthetic Disk Drive");            
        }

        internal static IEnumerable<ManagementObject> GetRessourcePools(this ManagementScope scope)
        {
            return scope.GetObjects("Msvm_ResourcePool");
        }

        internal static ManagementObject GetMemorySettingData(this ManagementObject virtualSystemSettingData)
        {
            return
                virtualSystemSettingData
                    .GetRelated("Msvm_MemorySettingData")
                        .OfType<ManagementObject>()
                        .FirstOrDefault();
        }

        internal static ManagementObject GetVirtualSystemSettingData(this ManagementObject system)
        {
            return
                system.GetRelated(
                    "Msvm_VirtualSystemSettingData",
                    "Msvm_SettingsDefineState",
                    null,
                    null,
                    "SettingData",
                    "ManagedElement",
                    false,
                    null)
                        .OfType<ManagementObject>()
                        .FirstOrDefault();
        }

        internal static IEnumerable<string> GetVirtualSwitchNames(this ManagementScope scope)
        {
            return scope.GetVirtualSwitches().Select(s => s["ElementName"].ToString());
        } 

        internal static IEnumerable<ManagementObject> GetVirtualSwitches(this ManagementScope scope)
        {
            return scope.GetObjects("Msvm_VirtualEthernetSwitch");
        }

        internal static ManagementObject GetVirtualSwitch(this ManagementScope scope, string switchName)
        {
            return
                scope.GetVirtualSwitches().FirstOrDefault(x => x["ElementName"].Equals(switchName));
        }

        internal static ManagementObject GetVirtualSystemSettingData(this ManagementScope scope, string virtualSystemIdentifier)
        {
            return scope.GetObjects("Msvm_VirtualSystemSettingData").FirstOrDefault(x => x["VirtualSystemIdentifier"] != null && x["VirtualSystemIdentifier"].Equals(virtualSystemIdentifier));
        }

        internal static ManagementObject GetVirtualSystemManagementService(this ManagementScope scope)
        {
            return scope.GetObject("MsVM_VirtualSystemManagementService");
        }

        internal static IEnumerable<ManagementObject> GetComputerSystems(this ManagementScope scope)
        {
            return scope.GetObjects("Msvm_ComputerSystem");
        }

        internal static ManagementObject GetVirtualMachine(this ManagementScope scope, string vmIdentity)
        {
            return scope.GetVirtualMachines().FirstOrDefault(x => x["name"].ToString().Equals(vmIdentity));
        } 

        internal static IEnumerable<ManagementObject> GetVirtualMachines(this ManagementScope scope)
        {
            return scope.GetComputerSystems().Where(x => x["Description"].ToString().ToLower().Contains("virtual machine"));
        }

        internal static ManagementObject GetObject(this ManagementScope scope, string serviceName)
        {
            return GetObjects(scope, serviceName).FirstOrDefault();
        }

        internal static IEnumerable<ManagementObject> GetObjects(this ManagementScope scope, string serviceName)
        {
            return new ManagementClass(scope, new ManagementPath(serviceName), null)
                .GetInstances()
                .OfType<ManagementObject>();
        }
    }
}
