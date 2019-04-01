using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ExcelAddIn.Objects;

namespace ExcelAddIn1
{
    public partial class ThisAddIn
    {
        public List<oCruce> _result = new List<oCruce>();
        public oCruce[] _TotalCruces = Assembler.LoadJson<oCruce[]>($"{ExcelAddIn.Access.Configuration.Path}\\jsons\\Cruces.json");
        public List<oCruce> _CrucesSinDiferencia = new List<oCruce>();
        public List<oCruce> _CrucesQueNoAplican = new List<oCruce>();
        public VerificacionDeCruce vdcUserControl;
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        public ControlImprimir control;
        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get { return myCustomTaskPane; }

        }
        public void Imprimir()
        {
            control = new ControlImprimir();
            myCustomTaskPane = this.CustomTaskPanes.Add(control, "Imprimir panel.xlsx");
            myCustomTaskPane.Width = 370;
            myCustomTaskPane.Visible = true;
        }
        public void CerrarImprimir()
        {
            this.CustomTaskPanes.Remove(myCustomTaskPane);
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Handler para instanciar el Control de usuario en cada instancia de Excel creada
            this.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_WorkbookActivate);

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
        void Application_SheetSelectionChange(object Sh, Excel.Range Target)
        {
            Globals.Ribbons.Ribbon2.btnAgregarIndice.Enabled = (!Target.AddressLocal.Contains(":"));
            Globals.Ribbons.Ribbon2.btnAgregarExplicacion.Enabled = (!Target.AddressLocal.Contains(":"));
            Globals.Ribbons.Ribbon2.btnEliminarIndice.Enabled = (!Target.AddressLocal.Contains(";"));// si  selecciona celdas intercaladas
            Globals.Ribbons.Ribbon2.btnEliminaeExplicacion.Enabled = (!Target.AddressLocal.Contains(";"));// si  selecciona celdas intercaladas
        }
        /// <summary>
        /// Método que crea un panel de validación de cruces por cada instancia de Excel
        /// </summary>
        /// <param name="wb"></param>
        private void Application_WorkbookActivate(Microsoft.Office.Interop.Excel.Workbook wb)
        {
            try
            {
                System.IO.FileInfo _ExcelFI = new System.IO.FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.Name);
                vdcUserControl = new VerificacionDeCruce();
                myCustomTaskPane = CustomTaskPanes.Add(vdcUserControl, "Verificación " + _ExcelFI.Name);
                myCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                myCustomTaskPane.DockPositionRestrict = Office.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
                myCustomTaskPane.Width = 515;
                myCustomTaskPane.Visible = false;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al instanciar el Task Panel de Validacion. [ThisAddIn].[Application_WorkbookActivate].[26]: {ex.Message}");
            }
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
