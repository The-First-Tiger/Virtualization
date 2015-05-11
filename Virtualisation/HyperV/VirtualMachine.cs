namespace Virtualization.HyperV
{
    public class VirtualMachine
    {
        public string VmIdentity { get; private set; }

        public string Name { get; set; }

        public VirtualMachineStates State { get; set; }

        public VirtualMachine(string vmIdentity)
        {
            this.VmIdentity = vmIdentity;
        }
    }
}
