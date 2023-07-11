using System;

namespace ExcelExportLibrary
{
    public class Program
    {
        /// <summary>
        ///     Function which is called via the Add-In. Starts the process for the export via Excel
        /// </summary>
        /// <param name="args"></param>
        public static void Main(string[] args)
        {
            // Attach to AssemblyResolver event and start tia portal
            AppDomain.CurrentDomain.AssemblyResolve += AssemblyResolver.OpennessLatestResolver;
            if (args.Length == 0)
                Run.StartExportToExcel();
            else
                Run.StartExportToExcel(args[0]);
        }
    }
}