using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class frmNewIndices : Form
    {
        public frmNewIndices()
        {
            InitializeComponent();
            txtCantIndices.Text = "1";
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            int cantRows = Convert.ToInt32(txtCantIndices.Text);
            Excel.Worksheet NewActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            InsertIndice(NewActiveWorksheet, cantRows, currentCell);
        }
        public void InsertIndice(Excel.Worksheet xlSht, int CantReg, Excel.Range currentCell)
        {

            int NroRow = currentCell.Row;
            int NroColumn = currentCell.Column;
            Excel.Range currentFind = null;         

            Excel.Range objRange = (Excel.Range)xlSht.Cells[NroRow, NroColumn];
            string IndicePrevio = objRange.get_Value(Type.Missing).ToString();
           

            int i = 1;
            Int64 indiceNvo = 0;
            while (i <= CantReg)
            {

                var rangej = xlSht.get_Range(string.Format("{0}:{0}", NroRow + i, Type.Missing));
                rangej.Select();

                rangej.Insert(Excel.XlInsertShiftDirection.xlShiftDown);


                indiceNvo = Convert.ToInt64(IndicePrevio) + 100;
                currentFind = objRange.Find(indiceNvo, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                                       Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                                        Type.Missing, Type.Missing);
                while (currentFind != null)
                {
                    indiceNvo = indiceNvo + 100;
                    currentFind = objRange.Find(indiceNvo, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                                      Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                                       Type.Missing, Type.Missing);

                }

                xlSht.Cells[NroRow + i, NroColumn] = "0" + Convert.ToString(indiceNvo);
                objRange = (Excel.Range)xlSht.Cells[NroRow + i, NroColumn];
                IndicePrevio = objRange.get_Value(Type.Missing).ToString();

                
                objRange.EntireRow.Font.Color = System.Drawing.Color.FromArgb(0, 0, 255);
                string targetRange = objRange.get_Address(
                Excel.XlReferenceStyle.xlA1);
                AddNamedRange(NroRow + i, NroColumn, "IA_" + "0" + Convert.ToString(indiceNvo));


                
                i = i + 1;

            }
           

        }
        public static void AddNamedRange(int row, int col, string myrango)
        {
            Microsoft.Office.Tools.Excel.NamedRange NamedRange1;

            Worksheet worksheet = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);


            Excel.Range cell = worksheet.Cells[row, col];
           
            NamedRange1 = worksheet.Controls.AddNamedRange(cell, myrango);

            NamedRange1.Sort(
                NamedRange1.Columns[1, Type.Missing], Excel.XlSortOrder.xlAscending,
                NamedRange1.Columns[2, Type.Missing], Type.Missing, Excel.XlSortOrder.xlAscending,
                Type.Missing, Excel.XlSortOrder.xlAscending,
                Excel.XlYesNoGuess.xlNo, Type.Missing, Type.Missing,
                Excel.XlSortOrientation.xlSortColumns,
                Excel.XlSortMethod.xlPinYin,
                Excel.XlSortDataOption.xlSortNormal,
                Excel.XlSortDataOption.xlSortNormal,
                Excel.XlSortDataOption.xlSortNormal);
           



        }
    }
}
