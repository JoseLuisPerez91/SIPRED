using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class Ribbon2
    {

        int NroFilaPrincipal = 0;
        int NroColPrincipal = 0;
        Excel.Range objRangeGlobal;

        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnNew_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btnNewIndice_Click(object sender, RibbonControlEventArgs e)
        {
            List<string> NombreRangos = new List<string>();
            Excel.Range objRange;
            string IndicePrevio = "";
            Excel.Worksheet ActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
            NroFilaPrincipal = currentCell.Row;
            NroColPrincipal = currentCell.Column;
            objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal, NroColPrincipal];
            IndicePrevio = objRange.get_Value(Type.Missing).ToString();
            objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal, 2];
            foreach (Excel.Name nm in wb.Names)
                NombreRangos.Add(nm.Name.ToString());

            if (!(NombreRangos.Contains("IA_" + IndicePrevio)))
                objRangeGlobal = null;

            string ConceptoPrevio = "";
            if (objRangeGlobal == null)
                ConceptoPrevio = objRange.get_Value(Type.Missing).ToString();


            if (objRangeGlobal == null)
            {

                objRangeGlobal = objRange;
            }
            else
            {
                objRange = objRangeGlobal;
                ConceptoPrevio = objRange.get_Value(Type.Missing).ToString();
            }


            if (ConceptoPrevio.Substring(0, 4).ToUpper() == "OTRO")
            {
                frmNewIndices NewIndices = new frmNewIndices();
                NewIndices.ShowDialog();
            }
            else
                MessageBox.Show("No es posible agregar índices debajo del índice " + IndicePrevio);


        }

        private void btnDelIndice_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;


            int NroRow = currentCell.Row;
            int NroColum = currentCell.Column;



            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = wb.ActiveSheet as Excel.Worksheet;
            List<string> NombreRangos = new List<string>();


            foreach (Excel.Name nm in wb.Names)
                NombreRangos.Add(nm.Name.ToString());

            Worksheet sheet = Globals.Factory.GetVstoObject(
               Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets[1]);

            Excel.Range objRange = (Excel.Range)ws.Cells[NroRow, NroColum];
            string IndiceActivo = objRange.get_Value(Type.Missing).ToString();
            int iNumeroFilas = ws.UsedRange.Rows.Count;
            int iNumeroColumnas = ws.UsedRange.Columns.Count;
            if (NombreRangos.Contains("IA_" + IndiceActivo))
            {
                objRange.EntireRow.Delete();

                sheet.Controls.Remove("IA_" + IndiceActivo);//remuevo el namesrange

                //  MessageBox.Show("Indice eliminado satisfactoriamente");
            }
            else
                MessageBox.Show("No puede eliminar indice base");
        }
    }
}
