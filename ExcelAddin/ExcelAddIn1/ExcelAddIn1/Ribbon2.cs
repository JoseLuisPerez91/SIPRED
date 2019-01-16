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
            objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal, 1];
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
                NroFilaPrincipal = objRange.Row;
                ConceptoPrevio = objRange.get_Value(Type.Missing).ToString();
            }


            if (ConceptoPrevio.Substring(0, 4).ToUpper() == "OTRO")
            {
               
                frmNewIndices NewIndices = new frmNewIndices(NroFilaPrincipal);
                
                NewIndices.ShowDialog();
            }
            else
                MessageBox.Show("No es posible agregar índices debajo del índice " + IndicePrevio);


        }

        public static void AddNamedRange(int row, int col, string myrango)
        {
            Microsoft.Office.Tools.Excel.NamedRange NamedRange1;

            Worksheet worksheet = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);


            Excel.Range cell = worksheet.Cells[row, col];

            NamedRange1 = worksheet.Controls.AddNamedRange(cell, myrango);



        }
        private void btnDelIndice_Click(object sender, RibbonControlEventArgs e)
        {

            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            List<string> NombreRangos = new List<string>();
            List<string> NombreRangosDEL = new List<string>();
            foreach (Excel.Name nm in wb.Names)
                NombreRangos.Add(nm.Name.ToString());



            Worksheet sheet = Globals.Factory.GetVstoObject(
               Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            Excel.Range objRange = null;
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;
            string IndiceActivo = "";
            bool Eliminar = false; ;
            try
            {
                Excel.Range objRangeI = (Excel.Range)sheet.Cells[NroFilaPrincipal, 1];
                long IndiceBase = Convert.ToInt64(objRangeI.get_Value(Type.Missing).ToString());
                
                int i = 0;
                foreach (Excel.Range cell in currentCell.Cells)
                {

                    try
                    {


                        objRange = (Excel.Range)sheet.Cells[cell.Row, cell.Column];
                        IndiceActivo = objRange.get_Value(Type.Missing).ToString();
                        //if (!NombreRangos.Contains("IA_" + IndiceActivo)) 
                        // IndiceBase = IndiceBase + ((NroFilaPrincipal+i) * 100);
                        //Indicebase = "0" + Convert.ToString(IndiceBase);
                        //  if (Indicebase!= IndiceActivo)
                        if (!NombreRangos.Contains("IA_" + IndiceActivo))
                        {
                            MessageBox.Show("No es posible eliminar un índice de formato guía");
                            Eliminar = false;
                            break;
                        }
                        else
                        {
                            Eliminar = true;
                            //  sheet.Controls.Remove();
                            NombreRangosDEL.Add("IA_" + IndiceActivo);
                        }
                        // i = i + 1;
                    }

                    catch
                    {

                        MessageBox.Show("NULL VALUE");

                    }

                }

                if (Eliminar)
                {
                    currentCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                    currentCell.EntireRow.Delete();
                    foreach (string Nm in NombreRangosDEL)
                        sheet.Controls.Remove(Nm);


                }
                i = 1;
                objRangeI.Select();
                objRange = (Excel.Range)sheet.Cells[NroFilaPrincipal+i, 1];
                IndiceActivo = objRange.get_Value(Type.Missing).ToString();
                long dif = 0;
                bool tienedif = false;
                while (NombreRangos.Contains("IA_"+IndiceActivo))
                {
                    dif = Convert.ToInt64(IndiceActivo) - IndiceBase;
                    while (dif != 100)
                    {
                        sheet.Controls.Remove("IA_" + IndiceActivo);
                        objRange.Value2 = "0" + Convert.ToString(Convert.ToInt64(IndiceActivo) - 100);
                        IndiceActivo = objRange.Value2;
                        dif = dif - 100;
                        tienedif = true;
                    }
                  if (tienedif)
                    AddNamedRange(objRange.Row, objRange.Column, "IA_"+objRange.Value2);

                    
                    IndiceBase = Convert.ToInt64(objRange.Value2);
                      //  dif = Convert.ToInt64(IndiceActivo) - IndiceBase;
                        i = i + 1;
                        objRange = (Excel.Range)sheet.Cells[NroFilaPrincipal + i, 1];
                    
                          IndiceActivo = objRange.get_Value(Type.Missing).ToString();
                      
                    
                }
            }
            catch
            {

                MessageBox.Show("NULL VALUE");

            }



            //REcalculo
        }
    }
}
