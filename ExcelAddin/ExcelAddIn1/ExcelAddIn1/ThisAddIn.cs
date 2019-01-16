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

            //----DATOS DE PRUEBA ...ALgerie
            Excel.Worksheet xlSht = Globals.ThisAddIn.Application.ActiveSheet;
            ((Excel.Range)xlSht.Cells[1, 1]).NumberFormat = "@";
            ((Excel.Range)xlSht.Cells[2, 1]).NumberFormat = "@";
            ((Excel.Range)xlSht.Cells[3, 1]).NumberFormat = "@";
            ((Excel.Range)xlSht.Cells[4, 1]).NumberFormat = "@";
            ((Excel.Range)xlSht.Cells[1, 1]).ColumnWidth = 20;
            ((Excel.Range)xlSht.Cells[2, 1]).ColumnWidth = 20;
            ((Excel.Range)xlSht.Cells[3, 1]).ColumnWidth = 20;
            ((Excel.Range)xlSht.Cells[4, 1]).ColumnWidth = 20;
            ((Excel.Range)xlSht.Cells[4, 1]).NumberFormat = "@";
            xlSht.Cells[1, 1] = "01080031000000";
            xlSht.Cells[1, 2] = "OTROS A FAVOR";
            xlSht.Cells[2, 1] = "01080032000000";
            xlSht.Cells[2, 2] = "EFECTO DE REEXPRESION";
            xlSht.Cells[3, 1] = "01080033000000";
            xlSht.Cells[3, 2] = "OTROS A CARGO";
            xlSht.Cells[4, 1] = "01080034000000";
            xlSht.Cells[4, 2] = "EFECTO DE REEXPRESION";
           
            RemoveCutCopyPasteMenuItems();//proteger

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            ResetCellMenu();
        }
        private Office.CommandBar GetCellContextMenu()
        {
             return this.Application.CommandBars["Cell"];
          
        }
        private void RemoveCutCopyPasteMenuItems()
        {
            Office.CommandBar contextMenu = GetCellContextMenu();

            for (int i = contextMenu.Controls.Count; i > 0; i--)
            {
                Office.CommandBarControl control = contextMenu.Controls[i];

                if ((control.Caption == "Cu&t") || (control.Caption == "Cor&tar") )control.Enabled=false; // Sample code: remove cut menu item 
                else if ((control.Caption == "&Copy") || (control.Caption == "&Copiar")) control.Enabled = false; // Sample code: remove copy menu item 
                else if ((control.accDescription.Contains("Paste"))  || (control.accDescription.Contains("Pegar"))
                    || (control.accDescription.Contains("Pe&gado")) || (control.accDescription.Contains("Insertar"))
                    || (control.accDescription.Contains("Eliminar"))
                    || (control.accDescription.Contains("Borrar"))) control.Enabled = false; // Sample code: remove any paste menu items 

            
            }

         
        }

       
        private void ResetCellMenu()
        {
            GetCellContextMenu().Reset(); // reset the cell context menu back to the default 
        }

        //protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        //{
        //    return new Ribbon1();
        //}


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
