using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Win32;

namespace ExcelExportLibrary
{
    public abstract class AssemblyResolver
    {
        private const string BasePath = "SOFTWARE\\Siemens\\Automation\\Openness\\";

        /// <summary>
        ///     Resolver for the openness assembly
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        /// <returns></returns>
        public static Assembly OpennessLatestResolver(object sender, ResolveEventArgs args)
        {
            var index = args.Name.IndexOf(',');
            if (index == -1) return null;
            var name = args.Name.Substring(0, index);


            var generalOpennessKey = Registry.LocalMachine.OpenSubKey(BasePath);
            var highestTiaEntry =
                generalOpennessKey.OpenSubKey(getHighestVersionName(generalOpennessKey) + "\\PublicAPI");
            var highestOpennessEntry = highestTiaEntry.OpenSubKey(getHighestVersionName(highestTiaEntry));


            if (highestOpennessEntry == null)
                return null;


            var oRegKeyValue = highestOpennessEntry.GetValue(name);
            if (oRegKeyValue != null)
            {
                var filePath = oRegKeyValue.ToString();
                var fullPath = Path.GetFullPath(filePath);
                if (File.Exists(fullPath)) return Assembly.LoadFrom(fullPath);
            }

            return null;
        }

        /// <summary>
        ///     Get the highest version which is available
        /// </summary>
        /// <param name="root"></param>
        /// <returns></returns>
        private static string getHighestVersionName(RegistryKey root)
        {
            var subKeys = root.GetSubKeyNames();

            var TiaVersions = subKeys.Select(key => (Key: key, Versioned: new Version(key)));
            var highest = TiaVersions.Max(entry => entry.Versioned);
            return TiaVersions.First(v => v.Versioned == highest).Key;
        }
    }
}