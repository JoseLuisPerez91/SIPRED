using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        private MyUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //myUserControl1 = new MyUserControl();
            //myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            //myCustomTaskPane.Visible = true;
            Globals.ThisAddIn.Application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
            Globals.ThisAddIn.Application.SheetSelectionChange += new Excel.AppEvents_SheetSelectionChangeEventHandler(Application_SheetSelectionChange);
            Globals.ThisAddIn.Application.SheetActivate += new Excel.AppEvents_SheetActivateEventHandler(app_SheetActivate);
        }
        private void app_SheetActivate(object sheet)
        {
            if (Globals.ThisAddIn.Application.DisplayAlerts)
            {
                Globals.Ribbons.Ribbon2.MensageBloqueo((Excel.Worksheet)sheet);
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        //protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        //{
        //    return new Ribbon1();
        //}


        void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            

            Globals.Ribbons.Ribbon2.btnAgregarIndice.Enabled = (!Target.AddressLocal.Contains(":"));
            Globals.Ribbons.Ribbon2.btnAgregarExplicacion.Enabled = (!Target.AddressLocal.Contains(":"));
            Globals.Ribbons.Ribbon2.btnEliminarIndice.Enabled = (!Target.AddressLocal.Contains(";"));// si  selecciona celdas intercaladas
            Globals.Ribbons.Ribbon2.btnEliminaeExplicacion.Enabled = (!Target.AddressLocal.Contains(";"));// si  selecciona celdas intercaladas

        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
