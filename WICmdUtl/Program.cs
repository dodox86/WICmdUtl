using System;
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
            "ProductCode", "Version", "Language", "ProductName"
        };

        private static string separator = "\t";

        private WindowsInstaller.Installer installer = null;

        static void Main(string[] args)
        {
            var program = new Program();
            program.Run();
        }

        private void Run()
        {
            installer = (WindowsInstaller.Installer)new Installer();

            OutputHeader();
            for (int i = 0; i < installer.Products.Count; i++)
            {
                OutputLine(i + 1, installer.Products[i]);
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
            for (int i = 1; i < PropertyNames.Length; i++)
            {
                sa[i] = installer.ProductInfo[productCode, PropertyNames[i]];
            }
            sa[1] = DecodeVersion(sa[1]);

            return string.Join(separator, sa);
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

