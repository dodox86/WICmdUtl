using System;
using System.Management;
using System.Collections.Generic;
using System.Diagnostics;

namespace WICmdUtl
{
    class Program
    {
        [System.Runtime.InteropServices.ComImport(),
        System.Runtime.InteropServices.Guid("000C1090-0000-0000-C000-000000000046")]
        class Installer { }

        // see Msi.h
        private static string[] PropertyNames = {
            "ProductCode", "UpgradeCode", "Version", "Language", "ProductName"
        };

        private static string separator = "\t";

        private WindowsInstaller.Installer installer = null;
        private ManagementObjectSearcher moSearcher = null;
        private Dictionary<string, string> updateCodeList = null;
        private List<string> options = new List<string>();

        private string targetProductCode = string.Empty;

        static void Main(string[] args)
        {
            var program = new Program();
            program.ParseCommandLine(args);
            program.Run();
        }

        private void ParseCommandLine(string[] args)
        {
            if (args.Length >= 1)
            {
                targetProductCode = args[0];
            }

            if (args.Length >= 2)
            {
                for (int i = 1; i < args.Length; i++)
                {
                    if (args[i].Length == 2)
                    {
                        if (args[i].Substring(0, 1) == "/")
                        {
                            options.Add(args[i].Substring(1, 1));
                        }
                    }
                }
            }
        }

        private void Run()
        {
            installer = (WindowsInstaller.Installer)new Installer();
            moSearcher = new ManagementObjectSearcher();
            moSearcher.Scope = new ManagementScope(@"\\.\ROOT\cimv2");

            if (options.Count == 0)
            {
                for (int i = 0; i < installer.Products.Count; i++)
                {
                    OutputLine(i + 1, installer.Products[i]);
                }
            }
            else
            {
                foreach (var option in options)
                {
                    switch (option)
                    {
                        case "u":
                            var upgradeCode = GetRelatedUpgradeCode(targetProductCode);
                            string s = string.Format("UpgradeCode: {0}", upgradeCode);
                            Console.WriteLine(s);
                            break;

                        default:
                            break;
                    }
                }
            }
        }

        private void OutputHeader()
        {
            Console.WriteLine(string.Join(separator, PropertyNames));
        }

        private void OutputLine(int no, string productCode)
        {
            var s = ToProductInfoLine(no, productCode);
            Debug.WriteLine(s);
            Console.WriteLine(s);
        }

        private string ToProductInfoLine(int no, string productCode)
        {
            string[] sa = new string[PropertyNames.Length];

            sa[0] = productCode;
            //sa[1] = GetRelatedUpgradeCode(productCode);
            sa[1] = string.Empty;
            for (int i = 2; i < PropertyNames.Length; i++)
            {
                sa[i] = installer.ProductInfo[productCode, PropertyNames[i]];
            }
            sa[3] = DecodeVersion(sa[1]);

            return string.Join(separator, sa);
        }

        private void BuildUpgradeCodeList()
        {
            updateCodeList.Clear();

            //string query = string.Format("SELECT * FROM Win32_Property WHERE Property='UpgradeCode'");
            ManagementScope scope = new ManagementScope(@"\\.\ROOT\cimv2");
            ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_Property");
            var searcher = new ManagementObjectSearcher(scope, query);
            ManagementObjectCollection cols = searcher.Get();

            foreach (ManagementObject mo in cols)
            {
                Debug.Assert(mo["ProductCode"] is string);
                Debug.Assert(mo["Value"] is string);
                updateCodeList.Add((string)mo["ProductCode"], (string)mo["UpgradeCode"]);
            }
        }

        private string GetRelatedUpgradeCode(string productCode)
        {
            string upgradeCode = string.Empty;
            string query = string.Format("SELECT Value FROM Win32_Property WHERE Property='UpgradeCode' AND ProductCode='{0}'", productCode);
            moSearcher.Query = new SelectQuery(query);
            ManagementObjectCollection cols = moSearcher.Get();
            foreach (ManagementObject mo in cols)
            {
                Debug.Assert(mo["Value"] is string);
                upgradeCode = (string)mo["Value"];
                Debug.WriteLine("UpgradeCode: " + upgradeCode);
                break;
            }
            cols.Dispose();

            return upgradeCode;
        }

        /// <summary>
        /// decodes version string to "major.minor.build"
        /// </summary>
        /// <param name="source">source string</param>
        /// <returns>human readable version string</returns>
        private static string DecodeVersion(string source)
        {
            string sresult = source;

            UInt32 result;
            if (UInt32.TryParse(source, out result))
            {
                // major,minor,build
                uint major = (result >> 24) & 0xff;
                uint minor = (result >> 16) & 0xff;
                uint build = result & 0xffff;

                sresult = string.Format("{0}.{1}.{2}", major, minor, build);
            }
            return sresult;
        }
    }
}

