using System;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using Siemens.Engineering.AddIn.Utilities;
using Siemens.Engineering.SW.Tags;

namespace ExportDataSetToExcel
{
    internal static class CliHandling
    {
        private const string ExecutablePath = "Delivery/ExcelExportLibrary.exe";

        /// <summary>
        ///     Prepares and starts the CLI process
        /// </summary>
        /// <param name="tagTable"></param>
        public static void RunExecutable(PlcTagTable tagTable)
        {
            var cliPath = GetOrExtractExecutable();
            var startInfo = new ProcessStartInfo
            {
                FileName = cliPath,
                CreateNoWindow = true,
                UseShellExecute = false
            };
            if (tagTable != null) startInfo.Arguments = "\"" + tagTable.Name + "\"";

            var cliProcess = new Process
            {
                EnableRaisingEvents = true,
                StartInfo = startInfo
            };
            cliProcess.Start();
        }

        /// <summary>
        ///     Gets the executable. If not available extract from .zip.
        /// </summary>
        /// <returns></returns>
        private static string GetOrExtractExecutable()
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string companyName = (asm.GetCustomAttribute(typeof(AssemblyCompanyAttribute)) as AssemblyCompanyAttribute)
                .Company;
            string applicationName = (asm.GetCustomAttribute(typeof(AssemblyTitleAttribute)) as AssemblyTitleAttribute)
                .Title;
            string baseFolder =
                Path.Combine(System.Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
                    companyName, "Automation", applicationName);

            if (!Directory.Exists(baseFolder))
            {
                Directory.CreateDirectory(baseFolder);
            }

            var targetPath = Path.Combine(baseFolder, ExecutablePath);

            if (File.Exists(targetPath))
            {
                if (CheckExistingFileIntegrity(targetPath))
                    return targetPath;
                else
                {
                    Directory.Delete(Path.GetDirectoryName(targetPath), true);
                }
            }

            var tmpZipPath = Path.Combine(baseFolder, "AddInContainedExe.zip");
            var assembly = Assembly.GetExecutingAssembly();
            using (var resource = assembly.GetManifestResourceStream("ExportDataSetToExcel.ExecutableDelivery.zip"))
            {
                using (var file = new FileStream(tmpZipPath, FileMode.Create, FileAccess.ReadWrite))
                {
                    resource.CopyTo(file);
                }
            }

            ZipFile.ExtractToDirectory(tmpZipPath, baseFolder);
            File.Delete(tmpZipPath);

            return targetPath;
        }

        /// <summary>
        ///     Check if the hash has changed
        /// </summary>
        /// <param name="existingFilePath"></param>
        /// <returns></returns>
        private static bool CheckExistingFileIntegrity(string existingFilePath)
        {
            var shaExistingFile = GetSha256(existingFilePath);

            var assembly = Assembly.GetExecutingAssembly();
            var shaCompare =
                new StreamReader(assembly.GetManifestResourceStream("ExportDataSetToExcel.ExecutableChecksum.txt"))
                    .ReadToEnd();
            shaCompare = shaCompare.Trim();

            return shaExistingFile.Equals(shaCompare, StringComparison.CurrentCultureIgnoreCase);
        }

        /// <summary>
        ///     Get the SHA256
        /// </summary>
        /// <param name="inputPath"></param>
        /// <returns></returns>
        private static string GetSha256(string inputPath)
        {
            using (var stream = File.OpenRead(inputPath))
            {
                var sha = new SHA256Managed();
                var checksum = sha.ComputeHash(stream);
                return BitConverter.ToString(checksum).Replace("-", string.Empty);
            }
        }
    }
}