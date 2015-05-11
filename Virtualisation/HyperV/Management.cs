namespace Virtualization.HyperV
{
    using System;
    using System.Collections.Generic;
    using System.Diagnostics;
    using System.Linq;
    using System.Management;
    using Extensions;

    public class Management
    {
        public string Host { get; private set; }

        public ConnectionOptions ConnectionOptions { get; private set; }

        public Management(string host = "localhost", string username = null, string password = null)
        {
            this.Host = host;

            if (!string.IsNullOrWhiteSpace(username)
                && !string.IsNullOrWhiteSpace(password))
            {
                this.ConnectionOptions = new ConnectionOptions()
                                         {
                                             Username = username,
                                             Password = password
                                         };
            }
        }

        public IEnumerable<VirtualMachine> GetVirtualMachines()
        {
            var scope = ManagementExtensions.GetScope(this.Host, this.ConnectionOptions);
            return scope.GetVirtualMachines()
                .Select(
                    vm =>
                        new VirtualMachine(vm["name"].ToString())
                        {
                            Name = vm["ElementName"].ToString(),
                            State = (VirtualMachineStates)(UInt16)vm["EnabledState"]
                        })
                .ToList();
        } 

        public VirtualMachine CreateVirtualMachine(string name)
        {
            var scope = ManagementExtensions.GetScope(this.Host, this.ConnectionOptions);
            var vmIdentity = scope.CreateVirtualMachine(name);

            return new VirtualMachine(vmIdentity) { Name = name, State = VirtualMachineStates.Disabled };
        }

        public void AddNetworkAdapter(VirtualMachine virtualMachine, string switchName)
        {
            var scope = ManagementExtensions.GetScope(this.Host, this.ConnectionOptions);
            scope.AddNetworAdapter(virtualMachine.VmIdentity, switchName);
        }

        public void SetDynamicRam(VirtualMachine virtualMachine, int ramSize = 4096, int? reservation = 1024, int? limit = 8192)
        {
            var scope = ManagementExtensions.GetScope(this.Host, this.ConnectionOptions);
            scope.SetDynamicRam(virtualMachine.VmIdentity, ramSize, reservation, limit);
        }

        public void SetVhd(VirtualMachine virtualMachine, string vhdPath)
        {
            var scope = ManagementExtensions.GetScope(this.Host, this.ConnectionOptions);
            scope.SetVhd(virtualMachine.VmIdentity, vhdPath);
        }

        public void StartVirtualMachine(VirtualMachine virtualMachine)
        {
            var scope = ManagementExtensions.GetScope(this.Host, this.ConnectionOptions);
            scope.StartVirtualMachine(virtualMachine.VmIdentity);
        }

        public IEnumerable<string> GetVirtualSwitches()
        {
            var scope = ManagementExtensions.GetScope(this.Host, this.ConnectionOptions);
            return scope.GetVirtualSwitchNames();
        } 
    }
}
