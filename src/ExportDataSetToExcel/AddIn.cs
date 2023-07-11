using System;
using System.Diagnostics;
using System.Linq;
using System.Security.Permissions;
using System.Windows.Forms;
using Siemens.Engineering;
using Siemens.Engineering.AddIn.Menu;
using Siemens.Engineering.AddIn.Permissions;
using Siemens.Engineering.HW;
using Siemens.Engineering.SW.Tags;

namespace ExportDataSetToExcel
{
    public class AddIn : ContextMenuAddIn
    {
        /// <summary>
        ///     The display name of the Add-In.
        /// </summary>
        private const string s_DisplayNameOfAddIn = "Export Tag Table Dataset To Excel";

        /// <summary>
        ///     The global TIA Portal Object
        ///     <para>It will be used in the TIA Add-In.</para>
        /// </summary>
        private TiaPortal _tiaPortal;

        /// <summary>
        ///     The constructor of the AddIn.
        ///     Creates an object of the class AddIn
        ///     Called from AddInProvider, when the first
        ///     right-click is performed in TIA
        ///     Motherclass' constructor of ContextMenuAddin
        ///     will be executed, too.
        /// </summary>
        /// <param name="tiaPortal">
        ///     Represents the actual used TIA Portal process.
        /// </param>
        public AddIn(TiaPortal tiaPortal) : base(s_DisplayNameOfAddIn)
        {
            /*
            * The acutal TIA Portal process is saved in the
            * global TIA Portal variable _tiaportal
            * tiaportal comes as input Parameter from the
            * AddInProvider
            */
            _tiaPortal = tiaPortal;
        }

        /// <summary>
        ///     The method is supplemented to include the Add-In
        ///     in the Context Menu of TIA Portal.
        ///     Called when a right-click is performed in TIA
        ///     and a mouse-over is performed on the name of the Add-In.
        /// </summary>
        /// <typeparam name="addInRootSubmenu">
        ///     The Add-In will be displayed in
        ///     the Context Menu of TIA Portal.
        /// </typeparam>
        /// <example>
        ///     ActionItems like Buttons/Checkboxes/Radiobuttons
        ///     are possible. In this example, only Buttons will be created
        ///     which will start the Add-In program code.
        /// </example>
        protected override void BuildContextMenuItems(ContextMenuAddInRoot addInRootSubmenu)
        {
            /* Method addInRootSubmenu.Items.AddActionItem
            * Will Create a Pushbutton with the text 'Start Add-In Code'
            * 1st input parameter of AddActionItem is the text of the
            * button
            * 2nd input parameter of AddActionItem is the clickDelegate,
            * which will be executed in case the button 'Start
            * Add-In Code' will be clicked/pressed.
            * 3rd input parameter of AddActionItem is the
            * updateStatusDelegate, which will be executed in
            * case there is a mouseover the button 'Start
            * Add-In Code'.
            * in <placeholder> the type of AddActionItem will be
            * specified, because AddActionItem is generic
            * AddActionItem<DeviceItem> will create a button that will be
            * displayed if a rightclick on a DeviceItem will be
            * performed in TIA Portal
            * AddActionItem<Project> will create a button that will be
            * displayed if a rightclick on the project name
            * will be performed in TIA Portal
            */
            addInRootSubmenu.Items.AddActionItem<DeviceItem>("Export To Excel", OnClick_Excel_PLC);
            addInRootSubmenu.Items.AddActionItem<PlcTagTable>("Export Table To Excel", OnClick_Excel_Table);
        }

        private void OnClick_Excel_PLC(MenuSelectionProvider<DeviceItem> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            try
            {
                DemandProcessStartPermission();
            }
            catch (Exception ex)
            {
                // Possible solution for exception handling. Line below is an example on how to log the messages.
                // _tiaPortal.GetService<FeedbackService>()?.Log(NotificationIcon.Error, ex.Message);
                MessageBox.Show("Error has occurred. " + ex.Message, "Excel Export", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            PlcTagTable tagTable = null;

            CliHandling.RunExecutable(tagTable);
        }


        private void OnClick_Excel_Table(MenuSelectionProvider<PlcTagTable> menuSelectionProvider)
        {
#if DEBUG
            Debugger.Launch();
#endif
            try
            {
                DemandProcessStartPermission();
            }
            catch (Exception ex)
            {
                // Possible solution for exception handling. Line below is an example on how to log the messages.
                // _tiaPortal.GetService<FeedbackService>()?.Log(NotificationIcon.Error, ex.Message);
                MessageBox.Show("Error has occurred. " + ex.Message, "Excel Export", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }

            var tagtable = menuSelectionProvider.GetSelection().First() as PlcTagTable;

            CliHandling.RunExecutable(tagtable);
        }

        private void DemandProcessStartPermission()
        {
            try
            {
                new ProcessStartPermission(PermissionState.Unrestricted).Demand();
            }
            catch (Exception ex)
            {
                // Possible solution for exception handling. Line below is an example on how to log the messages.
                // _tiaPortal.GetService<FeedbackService>()?.Log(NotificationIcon.Error, ex.Message);
                MessageBox.Show("Error has occurred. " + ex.Message, "Excel Export", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}