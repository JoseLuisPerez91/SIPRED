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
using Microsoft.Office.Core;


namespace ExcelAddIn1
{
    public partial class frmNewIndices : Form
    {
        int NroPrincipal = 0;
        public frmNewIndices(int NroFilaPrincipal)
        {
            InitializeComponent();
            txtCantIndices.Text = "1";
            NroPrincipal = NroFilaPrincipal;

        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            int cantRows = Convert.ToInt32(txtCantIndices.Text);
            Excel.Worksheet NewActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
            
            InsertIndice(NewActiveWorksheet, cantRows, currentCell);

            
            //NewActiveWorksheet.Application.CommandBars["Ply"].Controls["&Delete"].Enabled = false;
        }
        public void InsertIndice(Excel.Worksheet xlSht, int CantReg, Excel.Range currentCell)
        {
            Worksheet sheet = Globals.Factory.GetVstoObject(
             Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            int NroRow = currentCell.Row;
            int NroColumn = currentCell.Column;
            string IndicePrevio = "";
            long IndiceInicial = 0;
            bool tieneformula = false;
            Excel.Range currentFind = null;
            int iTotalColumns = 0;
            int k = 1;
            int i = 1;
            int t = 0;
            string[] targetRange = new string[CantReg];
            string[] targetRange1 = new string[CantReg];
            long indiceNvo = 0;
            currentCell = (Excel.Range)xlSht.Cells[NroRow, 1];
            IndicePrevio = currentCell.get_Value(Type.Missing).ToString();
            IndiceInicial = Convert.ToInt64(IndicePrevio);
            Excel.Range rangeall = currentCell.EntireRow;
            iTotalColumns = xlSht.UsedRange.Columns.Count;
            
            while (k<= iTotalColumns) { 
            
                if (rangeall.Cells[k].HasFormula)
                   tieneformula = true;

                k = k + 1;
            }
           

           
            while (i <= CantReg)
            {

               
                indiceNvo = Convert.ToInt64(IndicePrevio) + 100;
                currentFind = currentCell.Find(indiceNvo, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                                       Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                                        Type.Missing, Type.Missing);
               
                while (currentFind != null)
                {
                    NroRow = NroRow + 1;
                    indiceNvo = indiceNvo + 100;
                   // sheet.Controls.Remove("IA_0" + indiceNvo);
                    currentFind = currentCell.Find(indiceNvo, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                                      Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                                       Type.Missing, Type.Missing);
                  
                  //  AddNamedRange(NroRow + i, NroColumn, "IA_0" + Convert.ToString(indiceNvo));
                }

                var rangej = xlSht.get_Range(string.Format("{0}:{0}", NroRow + i, Type.Missing));
                rangej.Select();


                rangej.Insert(Excel.XlInsertShiftDirection.xlShiftDown);


                if (tieneformula)
                {
                    var rangeaCopy = xlSht.get_Range(string.Format("{0}:{0}", NroRow + i, Type.Missing));
                    iTotalColumns = xlSht.UsedRange.Columns.Count;
                    rangeall.Copy();
                    rangeaCopy.PasteSpecial(Excel.XlPasteType.xlPasteFormulas, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);

                   
                     k = 1;
                  while (k <= iTotalColumns)
                   { 
                       
                            if (!(rangeaCopy.Cells[k].HasFormula))
                               rangeaCopy.Cells[k].Value = "";

                            
                        k = k + 1;
                    }
                }

                xlSht.Cells[NroRow + i, 1] = "0" + Convert.ToString(indiceNvo);
                

                currentCell = (Excel.Range)xlSht.Cells[NroRow + i, 1];
                IndicePrevio = currentCell.get_Value(Type.Missing).ToString();


                currentCell = xlSht.Range[xlSht.Cells[NroRow + i, 1], xlSht.Cells[NroRow + i, 3]];
                currentCell.Font.Color = System.Drawing.Color.FromArgb(0, 0, 255);

                //currentCell = (Excel.Range)xlSht.Cells[NroRow + i, 3];
                //targetRange[i-1]= currentCell.get_Address(Excel.XlReferenceStyle.xlA1);


                //currentCell = (Excel.Range)xlSht.Cells[NroRow + i, 4];
                //targetRange1[i-1] = currentCell.get_Address(Excel.XlReferenceStyle.xlA1);

               

                AddNamedRange(NroRow + i, NroColumn, "IA_" + Convert.ToString(IndicePrevio));
                
                i = i + 1;

             

            }
            
            Excel.Range Sum_Range = xlSht.get_Range("C" + (NroPrincipal).ToString(), "C" + (NroPrincipal).ToString());
           
            Sum_Range.Formula = "=sum(C" + (NroPrincipal + 1).ToString() + ":C" + (NroRow + CantReg).ToString();

            Sum_Range = xlSht.get_Range("D" + (NroPrincipal).ToString(), "D" + (NroPrincipal).ToString());

            Sum_Range.Formula = "=sum(D" + (NroPrincipal + 1).ToString() + ":D" + (NroRow + CantReg).ToString();


        }
        public static void AddNamedRange(int row, int col, string myrango)
        {
            Microsoft.Office.Tools.Excel.NamedRange NamedRange1;

            Worksheet worksheet = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);


            Excel.Range cell = worksheet.Cells[row, col];
           
            NamedRange1 = worksheet.Controls.AddNamedRange(cell, myrango);

            //NamedRange1.Sort(
            //    NamedRange1.Columns[1, Type.Missing], Excel.XlSortOrder.xlAscending,
            //    NamedRange1.Columns[2, Type.Missing], Type.Missing, Excel.XlSortOrder.xlAscending,
            //    Type.Missing, Excel.XlSortOrder.xlAscending,
            //    Excel.XlYesNoGuess.xlNo, Type.Missing, Type.Missing,
            //    Excel.XlSortOrientation.xlSortColumns,
            //    Excel.XlSortMethod.xlPinYin,
            //    Excel.XlSortDataOption.xlSortNormal,
            //    Excel.XlSortDataOption.xlSortNormal,
            //    Excel.XlSortDataOption.xlSortNormal);


        }
    }
}
