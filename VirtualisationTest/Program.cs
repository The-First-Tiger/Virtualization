using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VirtualisationTest
{
    using Virtualization.HyperV;

    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Create VM: ");
            var input = Console.ReadLine();

            Console.Write("RAM [MB]: ");
            var ram = Convert.ToInt32(Console.ReadLine());

            Console.Write("VHD-Path: ");
            var vhdPath = Console.ReadLine();

            Console.Write("Virtual Switch Name: ");
            var switchName = Console.ReadLine();

            var management = new Management();
            var virtualMachine = management.CreateVirtualMachine(input);

            management.SetDynamicRam(virtualMachine, ram);
            management.SetVhd(virtualMachine, vhdPath);
            management.AddNetworkAdapter(virtualMachine, switchName);

            management.StartVirtualMachine(virtualMachine);            

            Console.WriteLine("VM {0} created.", input);
            Console.ReadKey();
        }
    }
}
