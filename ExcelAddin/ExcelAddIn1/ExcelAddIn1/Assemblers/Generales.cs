using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.Diagnostics;
using ExcelAddIn.Objects;

namespace ExcelAddIn1
{
    public class Generales
    {

        public static void InsertIndice(Excel.Worksheet xlSht, int CantReg, Excel.Range currentCell, bool ConFormula, int NroPrincipal)
        {
            Worksheet sheet = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Range currentFind = null;
            Excel.Range currentFindExpl = null;
            int NroRow = currentCell.Row;
            int NroColumn = currentCell.Column;
            string IndicePrevio = "";
            long IndiceInicial = 0;
           
            int iTotalColumns = 0;
            int k = 1;
            int i = 1;
            long indiceNvo = 0;
            int CantExpl = 0;
            currentCell = (Excel.Range)xlSht.Cells[NroRow, 1];
            IndicePrevio = currentCell.get_Value(Type.Missing).ToString();
            currentFindExpl = (Excel.Range)xlSht.Cells[NroRow + 1, 1];
            if (currentFindExpl.get_Value(Type.Missing) != null)
                if (currentFindExpl.get_Value(Type.Missing).ToString().ToUpper().Trim() == "EXPLICACION")
                    NroRow++;
            IndiceInicial = Convert.ToInt64(IndicePrevio) + 100;
            int rowexpl = 0;
            List<int> FilasExplicacion = new List<int>();
            int CantRango = 0; long IndiceInicialx = IndiceInicial;
            foreach (Excel.Name cname in Globals.ThisAddIn.Application.Names)
            {

                if (cname.Name == "IA_0" + Convert.ToString(IndiceInicialx))
                {


                    CantRango++;

                    IndiceInicialx = IndiceInicialx + 100;

                    rowexpl = cname.RefersToRange.Cells.Row + 1;
                    currentFindExpl = (Excel.Range)xlSht.Cells[rowexpl, 1];
                    if (currentFindExpl.get_Value(Type.Missing) != null)
                        if (currentFindExpl.get_Value(Type.Missing).ToString().ToUpper().Trim() == "EXPLICACION")
                        {
                            CantExpl++;
                            if (!FilasExplicacion.Contains(rowexpl + CantReg))//los indices que tienen explicacion la fila actual + los registros que ingresó nvos
                                FilasExplicacion.Add(rowexpl + CantReg);
                        }

                }


            }


            currentFind = currentCell.Find(IndiceInicial, Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                                     Type.Missing, Type.Missing);





            int NroRowx = 0;

            CantRango = CantRango + CantExpl;
            int NroPrincipalAux = DameRangoPrincipal(NroPrincipal, xlSht);
            while (i <= CantReg)
            {


                indiceNvo = Convert.ToInt64(IndicePrevio) + 100;
               
                

                Excel.Range rangej = xlSht.get_Range(string.Format("{0}:{0}", NroRow + i, Type.Missing));


                rangej.Select();


                rangej.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);


              
                if (ConFormula)
                {
                    var rangeall = xlSht.get_Range(string.Format("{0}:{0}", NroPrincipalAux - 1, Type.Missing));
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
                sheet.Controls.Remove("IA_0" + indiceNvo);
                AddNamedRange(NroRow + i, 1, "IA_0" + Convert.ToString(indiceNvo));

                currentCell = (Excel.Range)xlSht.Cells[NroRow + i, 1];
                IndicePrevio = currentCell.get_Value(Type.Missing).ToString();


                currentCell = xlSht.Range[xlSht.Cells[NroRow + i, 1], xlSht.Cells[NroRow + i, 3]];
                currentCell.Font.Color = System.Drawing.Color.FromArgb(0, 0, 255);

                ((Excel.Range)xlSht.Cells[NroRow + 1, 2]).NumberFormat = "@"; // le doy formato text al concepto
                ((Excel.Range)xlSht.Cells[NroRow + 1, 2]).WrapText = true;
                

                i++;



            }



            if (currentFind != null)
            {


                NroRowx = NroRow + CantReg;
                currentCell = (Excel.Range)xlSht.Cells[NroRowx, 1];
                IndicePrevio = currentCell.get_Value(Type.Missing).ToString();

                int j = 1;


                while (j <= CantRango)
                {


                    if (!FilasExplicacion.Contains(NroRowx + j))
                    {
                        //  NroRowx++;
                        
                        indiceNvo = Convert.ToInt64(IndicePrevio) + 100;

                        xlSht.Cells[NroRowx + j, 1] = "0" + Convert.ToString(indiceNvo);
                        sheet.Controls.Remove("IA_0" + indiceNvo);
                        AddNamedRange(NroRowx + j, 1, "IA_0" + Convert.ToString(indiceNvo));

                        currentCell = (Excel.Range)xlSht.Cells[NroRowx + j, 1];
                        IndicePrevio = currentCell.get_Value(Type.Missing).ToString();
                    }
                    j++;
                }


            }

            
           
            List<oSubtotal> ColumnasST = DameColumnasST();
            Excel.Range Sum_Range = null;
            int NroFinal = NroRow + CantReg + CantRango;

            foreach (oSubtotal ST in ColumnasST)
            {
                if (ST.Hoja == xlSht.Name)
                {
                    Sum_Range = xlSht.get_Range(ST.Columna + (NroPrincipalAux).ToString(), ST.Columna + (NroPrincipalAux).ToString());
                    
                    Sum_Range.Formula = "=sum(" + ST.Columna + (NroPrincipalAux + 1).ToString() + ":" + ST.Columna + (NroFinal).ToString();
                }
            }

            
            Sum_Range = xlSht.get_Range("B" + (NroPrincipal).ToString(), "B" + (NroPrincipal).ToString());

            Sum_Range.Select();

            

        }

        public static void InsertaExplicacion(Excel.Worksheet xlSht, Excel.Range currentCell, string Explicacion)
        {

            var rangej = xlSht.get_Range(string.Format("{0}:{0}", currentCell.Row + 1, Type.Missing));
            rangej.Select();


            rangej.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            xlSht.Cells[currentCell.Row + 1, 1] = " EXPLICACION ";
            ((Excel.Range)xlSht.Cells[currentCell.Row + 1, 2]).NumberFormat = "@";
            ((Excel.Range)xlSht.Cells[currentCell.Row + 1, 2]).WrapText = true;
            xlSht.Cells[currentCell.Row + 1, 2] = Explicacion;


            currentCell.Select();

            currentCell = xlSht.Range[xlSht.Cells[currentCell.Row + 1, 1], xlSht.Cells[currentCell.Row + 1, 2]];
            currentCell.Font.Color = System.Drawing.Color.FromArgb(0, 0, 255);
          

        }

        public static void AddNamedRange(int row, int col, string myrango)
        {
            Microsoft.Office.Tools.Excel.NamedRange NamedRange1;

            Worksheet worksheet = Globals.Factory.GetVstoObject(
                Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);


            Excel.Range cell = worksheet.Cells[row, col];
            try
            {
                NamedRange1 = worksheet.Controls.AddNamedRange(cell, myrango);
            }
            catch
            {

            }



        }

        public static int DameRangoPrincipal(int NroPrincipal, Excel.Worksheet xlSht)
        {
            int NroPrincipalAux = NroPrincipal;
            Excel.Range objRange = (Excel.Range)xlSht.Cells[NroPrincipal, 2];
            var ConceptoPrevio = objRange.get_Value(Type.Missing);
            if (ConceptoPrevio != null)
            {
                ConceptoPrevio = ConceptoPrevio.ToString();

                if  (EsConceptoValido(ConceptoPrevio))
                {
                    

                        while (NroPrincipalAux > 0)
                        {
                            objRange = (Excel.Range)xlSht.Cells[NroPrincipalAux, 2];
                            ConceptoPrevio = objRange.get_Value(Type.Missing);
                            if (ConceptoPrevio != null)
                            {
                                ConceptoPrevio = ConceptoPrevio.ToString();
                           
                               if (EsConceptoValido(ConceptoPrevio))
                                   break;
                            }
                            NroPrincipalAux--;
                        }
                    
                }
                else
                {
                    while (NroPrincipalAux > 0)
                    {
                        objRange = (Excel.Range)xlSht.Cells[NroPrincipalAux, 2];
                        ConceptoPrevio = objRange.get_Value(Type.Missing);
                        if (ConceptoPrevio != null)
                        {
                            ConceptoPrevio = ConceptoPrevio.ToString();
                            if (EsConceptoValido(ConceptoPrevio))
                                break;
                           
                        }
                        NroPrincipalAux--;
                    }

                }
            }
            else
            {
                while (NroPrincipalAux > 0)
                {
                    objRange = (Excel.Range)xlSht.Cells[NroPrincipalAux, 2];
                    ConceptoPrevio = objRange.get_Value(Type.Missing);
                    if (ConceptoPrevio != null)
                    {
                        ConceptoPrevio = ConceptoPrevio.ToString();
                        if (EsConceptoValido(ConceptoPrevio))
                            break;
                        
                    }
                    NroPrincipalAux--;
                }

            }


            return NroPrincipalAux;
        }

        public static bool EsConceptoValido(string Concepto)
        {
            bool CncValido = false;
            List<oConcepto> ConceptVal = new List<oConcepto>();
            ConceptVal = DameConceptosValidos();

            foreach (oConcepto c in ConceptVal)
            {
                if (Concepto.Length >= c.Caracteres)
                {
                    if (Concepto.ToUpper().Substring(0, c.Caracteres).Contains(c.Descripcion.ToUpper()))
                    {
                        CncValido = true;
                        break;
                    }
                }

            }
            return CncValido;
        }
        public static List<oSubtotal> DameColumnasST()
        {
            List<oSubtotal> Subtotales = new List<oSubtotal>();
          
            string _Path = ExcelAddIn.Access.Configuration.Path;

            oRootobject _Root = Assembler.LoadJson<oRootobject>($"{_Path}\\jsons\\Indices.json");
           
            Subtotales = _Root.Subtotales;
           
            return Subtotales;
        }

        public static List<oConcepto> DameConceptosValidos()
        {
            List<oConcepto> Conceptos = new List<oConcepto>();
            string _Path = ExcelAddIn.Access.Configuration.Path;

            oRootobject _Root = Assembler.LoadJson<oRootobject>($"{_Path}\\jsons\\Indices.json");

           
            Conceptos = _Root.Conceptos;
          
            return Conceptos;
        }

    }

    

 
    
}
