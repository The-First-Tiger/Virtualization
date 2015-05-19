namespace Virtualization.HyperV
{
    using System;
    using System.Management;
    using Extensions;

    internal class VirtualSystemManagementService
    {
        private const int Successful = 0;

        private readonly ManagementObject virtualSystemManagementService;

        public VirtualSystemManagementService(ManagementScope scope)
        {
            this.virtualSystemManagementService = scope.GetVirtualSystemManagementService();
        }

        public ManagementBaseObject DefineSystem()
        {
            var definedSystem = this.virtualSystemManagementService.InvokeMethod(
                "DefineSystem",
                this.virtualSystemManagementService.GetMethodParameters("DefineSystem"),
                null);

            if (definedSystem == null)
            {
                throw new InvalidOperationException("DefineSystem failed");
            }

            var returnValue = (uint)definedSystem["returnvalue"];

            if (returnValue != Successful)
            {
                throw new InvalidOperationException("DefineSystem failed");
            }

            var vmPath = definedSystem["ResultingSystem"] as string;

            return new ManagementObject(vmPath);
        }
    }
}
