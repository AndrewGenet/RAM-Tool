using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;
using System.IO;

namespace win32_physicalMemoryTool
{
    public partial class Form1 : Form
    {

        //TextWriter class
        TextWriter _writer = null;

        public Form1()
        {
            InitializeComponent();
        }

        int slotsUsed = 0;

        // Retrieving Physical Ram Memory
        public static string GetPhysicalMemory()
        {
            ManagementScope oMs = new ManagementScope();
            ObjectQuery oQuery = new ObjectQuery("SELECT Capacity FROM Win32_PhysicalMemory");
            ManagementObjectSearcher oSearcher = new ManagementObjectSearcher(oMs, oQuery);
            ManagementObjectCollection oCollection = oSearcher.Get();

            long MemSize = 0;
            long mCap = 0;

            // In case more than one Memory sticks are installed
            foreach (ManagementObject obj in oCollection)
            {
                mCap = Convert.ToInt64(obj["Capacity"]);
                MemSize += mCap;
            }
            MemSize = ((MemSize / 1024) / 1024) / 1024;
            return MemSize.ToString() + " GB";
        }

        // Retrieving No of Ram Slot on Motherboard
        public static string GetNoRamSlots()
        {

            int MemSlots = 0;
            ManagementScope oMs = new ManagementScope();
            ObjectQuery oQuery2 = new ObjectQuery("SELECT MemoryDevices FROM Win32_PhysicalMemoryArray");
            ManagementObjectSearcher oSearcher2 = new ManagementObjectSearcher(oMs, oQuery2);
            ManagementObjectCollection oCollection2 = oSearcher2.Get();
            foreach (ManagementObject obj in oCollection2)
            {
                MemSlots = Convert.ToInt32(obj["MemoryDevices"]);

            }
            return MemSlots.ToString();
        }

        //Operating System
        public static string GetOSInformation()
        {
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem");
            foreach (ManagementObject wmi in searcher.Get())
            {
                try
                {
                    return ((string)wmi["Caption"]).Trim() + Environment.NewLine + "Version: " + (string)wmi["Version"] + ", " + (string)wmi["OSArchitecture"];
                }
                catch { }
            }
            return "BIOS Maker: Unknown";
        }

        //Computer Name
        public static String GetComputerName()
        {
            ManagementClass mc = new ManagementClass("Win32_ComputerSystem");
            ManagementObjectCollection moc = mc.GetInstances();
            String info = String.Empty;
            foreach (ManagementObject mo in moc)
            {
                info = (string)mo["Name"];
            }
            return info;
        }



        private void Form1_Load(object sender, EventArgs e)
        {

            // Instantiate the writer
            _writer = new TextBoxStreamWriter(txtConsole);
            // Redirect the out Console stream
            Console.SetOut(_writer);

            ManagementObjectSearcher searcher2 =
                        new ManagementObjectSearcher("root\\CIMV2",
                        "SELECT * FROM Win32_PhysicalMemory");

                foreach (ManagementObject queryObj in searcher2.Get())
                {
                    slotsUsed = slotsUsed + 1;
                    Console.WriteLine("-----------------------------------");
                    Console.WriteLine("Physical Memory instance");
                    Console.WriteLine("-----------------------------------");
                    //Console.WriteLine("BankLabel: {0}", queryObj["BankLabel"]);
                    Console.WriteLine("Capacity: {0}", queryObj["Capacity"]);
                    //Console.WriteLine("Caption: {0}", queryObj["Caption"]);
                    //Console.WriteLine("CreationClassName: {0}", queryObj["CreationClassName"]);
                    //Console.WriteLine("DataWidth: {0}", queryObj["DataWidth"]);
                    //Console.WriteLine("Description: {0}", queryObj["Description"]);
                    Console.WriteLine("DeviceLocator: {0}", queryObj["DeviceLocator"]);
                    //Console.WriteLine("FormFactor: {0}", queryObj["FormFactor"]);
                    //Console.WriteLine("HotSwappable: {0}", queryObj["HotSwappable"]);
                    //Console.WriteLine("InstallDate: {0}", queryObj["InstallDate"]);
                    Console.WriteLine("InterleaveDataDepth: {0}", queryObj["InterleaveDataDepth"]);
                    Console.WriteLine("InterleavePosition: {0}", queryObj["InterleavePosition"]);
                    Console.WriteLine("Manufacturer: {0}", queryObj["Manufacturer"]);
                    Console.WriteLine("MemoryType: {0}", queryObj["MemoryType"]);
                    //Console.WriteLine("Model: {0}", queryObj["Model"]);
                    //Console.WriteLine("Name: {0}", queryObj["Name"]);
                    //Console.WriteLine("OtherIdentifyingInfo: {0}", queryObj["OtherIdentifyingInfo"]);
                    Console.WriteLine("PartNumber: {0}", queryObj["PartNumber"]);
                    Console.WriteLine("PositionInRow: {0}", queryObj["PositionInRow"]);
                    //Console.WriteLine("PoweredOn: {0}", queryObj["PoweredOn"]);
                    //Console.WriteLine("Removable: {0}", queryObj["Removable"]);
                    //Console.WriteLine("Replaceable: {0}", queryObj["Replaceable"]);
                    //Console.WriteLine("SerialNumber: {0}", queryObj["SerialNumber"]);
                    //Console.WriteLine("SKU: {0}", queryObj["SKU"]);
                    Console.WriteLine("Speed: {0}", queryObj["Speed"]);
                    //Console.WriteLine("Status: {0}", queryObj["Status"]);
                    //Console.WriteLine("Tag: {0}", queryObj["Tag"]);
                    //Console.WriteLine("TotalWidth: {0}", queryObj["TotalWidth"]);
                    //Console.WriteLine("TypeDetail: {0}", queryObj["TypeDetail"]);
                    //Console.WriteLine("Version: {0}", queryObj["Version"]);

                    groupBox1.Text = "Showing Information for: " + GetComputerName();
                    label1.Text = GetOSInformation();
                    label2.Text = "Total RAM Slots: " + GetNoRamSlots();
                    label3.Text = "Slots In Use: " + Convert.ToString(slotsUsed);
                    label4.Text = "Total RAM Installed: " + GetPhysicalMemory();
                } 
            }
        }
    }

