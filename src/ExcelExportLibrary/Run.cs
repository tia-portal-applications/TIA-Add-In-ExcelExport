using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Siemens.Engineering;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Tags;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace ExcelExportLibrary
{
    public static class Run
    {
        public static TiaPortal _tiaPortal;

        /// <summary>
        ///     Starts the export if no TagTable is given
        /// </summary>
        public static void StartExportToExcel()
        {
            _tiaPortal = TiaPortal.GetProcesses()[0].Attach();
            var tagTableModelList = PrepareDataSetOfTagTable();

            ExportDataToExcel(tagTableModelList);
        }

        /// <summary>
        ///     Starts the export with an explicit TagTable
        /// </summary>
        /// <param name="tagtablename"></param>
        public static void StartExportToExcel(string tagtablename)
        {
            _tiaPortal = TiaPortal.GetProcesses()[0].Attach();
            var tagtableModelList = PrepareDataSetOfTagTable(tagtablename);
            ExportDataToExcel(tagtableModelList);
        }

        /// <summary>
        ///     Prepares the export of the TagTable without an explicit table
        /// </summary>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static List<TagTableModel> PrepareDataSetOfTagTable()
        {
            // Iteare over devices until any PLC found
            var plcSwTarget = FetchPLCSWTarget();

            var tagTableModelList = new List<TagTableModel>();

            // If no PLC found on the project,then warn the user to add at least one PLC
            if (plcSwTarget == null)
                throw new Exception(
                    "To export tag table from project tree, project should have at least one PLC device!");

            var tagTableComposition = ((PlcSoftware)plcSwTarget).TagTableGroup.TagTables;

            foreach (var tagTable in tagTableComposition)
            {
                if (tagTable.IsDefault)
                    continue;
                SetTagTableModel(tagTable, tagTableModelList);
            }

            return tagTableModelList;
        }

        /// <summary>
        ///     Prepares the export of the TagTable with an explicit table
        /// </summary>
        /// <param name="tablename"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        private static List<TagTableModel> PrepareDataSetOfTagTable(string tablename)
        {
            // Iteare over devices until any PLC found
            var plcSwTarget = FetchPLCSWTarget();

            var tagTableModelList = new List<TagTableModel>();

            // If no PLC found on the project,then warn the user to add at least one PLC
            if (plcSwTarget == null)
                throw new Exception(
                    "To export tag table from project tree, project should have at least one PLC device!");

            var tagTableComposition = ((PlcSoftware)plcSwTarget).TagTableGroup.TagTables;

            foreach (var tagTable in tagTableComposition)
            {
                if (tagTable.Name != tablename)
                    continue;
                SetTagTableModel(tagTable, tagTableModelList);
            }

            return tagTableModelList;
        }

        /// <summary>
        ///     Set model of TagTable
        /// </summary>
        /// <param name="tagTable"></param>
        /// <param name="tagTableModelList"></param>
        private static void SetTagTableModel(PlcTagTable tagTable, List<TagTableModel> tagTableModelList)
        {
            foreach (var tag in tagTable.Tags)
            {
                var tagTableModel = new TagTableModel();

                tagTableModel.TagName = tag.Name;
                tagTableModel.DataType = tag.DataTypeName;
                tagTableModel.Address = tag.LogicalAddress;
                tagTableModel.ExternalAccessible = tag.ExternalAccessible;
                tagTableModel.ExternalVisible = tag.ExternalVisible;
                tagTableModel.ExternalWritable = tag.ExternalWritable;

                tagTableModelList.Add(tagTableModel);
            }
        }

        /// <summary>
        ///     Fetch the PLC software target
        /// </summary>
        /// <returns></returns>
        private static Software FetchPLCSWTarget()
        {
            Software plcSwTarget = null;


            foreach (var device in GetDevices())
            foreach (var item in device.DeviceItems)
            {
                plcSwTarget = item.GetService<SoftwareContainer>()?.Software;

                if (plcSwTarget != null &&
                    plcSwTarget is PlcSoftware)
                    return
                        plcSwTarget; //If not null and type of PlcSoftware, then we can exit since we have found our target
            }

            return null;
        }

        /// <summary>
        /// Returns the Devicecomposition. Either from the LocalSessions or the Project
        /// </summary>
        /// <returns></returns>
        private static DeviceComposition GetDevices()
        {
            return _tiaPortal.Projects.Count != 0 ? _tiaPortal.Projects[0].Devices :
                _tiaPortal.LocalSessions.Count != 0 ? _tiaPortal.LocalSessions[0].Project.Devices : null;
        }

        #region Excel Interop 16.0

        /// <summary>
        ///     Exports the TagTable into a excel table
        /// </summary>
        /// <param name="tagTableModelList"></param>
        private static void ExportDataToExcel(List<TagTableModel> tagTableModelList)
        {
            //Check file exists
            var folderPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            var fileFullname = Path.Combine(folderPath, "Output.xlsx");
            if (File.Exists(fileFullname))
                try
                {
                    File.Delete(fileFullname);
                }
                catch (IOException)
                {
                    MessageBox.Show("Close Output.xlsx and try again",
                        "Excel Export",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

            // Start Excel and get Application object.  
            var excel = new Application();

            var excelworkBook = excel.Workbooks.Add(Type.Missing);


            // Workk sheet  
            var excelSheet = (Worksheet)excelworkBook.ActiveSheet;
            excelSheet.Name = "Sheet-1";

            excelSheet.Cells[1, 1] = "Name";
            excelSheet.Cells[1, 2] = "Data Type";
            excelSheet.Cells[1, 3] = "Address";
            excelSheet.Cells[1, 4] = "External Accessible";
            excelSheet.Cells[1, 5] = "External Visible";
            excelSheet.Cells[1, 6] = "External Writable";

            excelSheet.Cells.Font.Color = Color.Black;

            var rowCount = 1;
            var columnCount = 1;

            foreach (var data in tagTableModelList)
            {
                rowCount++;

                excelSheet.Cells[rowCount, columnCount] = data.TagName;
                excelSheet.Cells[rowCount, columnCount + 1] = data.DataType;
                excelSheet.Cells[rowCount, columnCount + 2] = data.Address;
                excelSheet.Cells[rowCount, columnCount + 3] = data.ExternalAccessible;
                excelSheet.Cells[rowCount, columnCount + 4] = data.ExternalVisible;
                excelSheet.Cells[rowCount, columnCount + 5] = data.ExternalWritable;
            }

            var excelCellrange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[rowCount, 6]];
            excelCellrange.EntireColumn.AutoFit();
            var border = excelCellrange.Borders;
            border.LineStyle = XlLineStyle.xlContinuous;
            border.Weight = 2d;


            excelworkBook.SaveAs(fileFullname);
            excelworkBook.Close();
            excel.Quit();
            //Inform user that export was successful
            MessageBox.Show("Successfully exported Tagtablecontent",
                "Excel Export",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        #endregion
    }
}