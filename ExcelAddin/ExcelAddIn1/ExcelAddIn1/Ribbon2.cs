using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using ExcelAddIn.Objects;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.IO;

namespace ExcelAddIn1 {
    public partial class Ribbon2 {
        int NroFilaPrincipal = 0;
        int NroColPrincipal = 0;
        bool tieneformula = false;

        private void Ribbon2_Load(object sender, RibbonUIEventArgs e) {

        }

        private void btnNew_Click(object sender, RibbonControlEventArgs e) {
            Nuevo _New = new Nuevo();
            _New.Show();
        }

        private void btnCruces_Click(object sender, RibbonControlEventArgs e)
        {
            Cruce _Cruce = new Cruce();
            _Cruce.Show();
        }

        private void btnPlantilla_Click(object sender, RibbonControlEventArgs e)
        {
            LoadTemplate _Template = new LoadTemplate();
            _Template.Show();
        }

        private void btnAgregarIndice_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet ActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Range objRange;
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
            NroFilaPrincipal = currentCell.Row;
            NroColPrincipal = currentCell.Column;
            int iTotalColumns; int k = 1;
            bool puedeinsertar = false;
            string IndicePrevio;

            string tag = "";
            try
            {

                objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal, 1];
                IndicePrevio = objRange.get_Value(Type.Missing).ToString();
                if (IndicePrevio.ToUpper().Trim() != "EXPLICACION")
                {

                    

                    foreach (Excel.Name item in wb.Names)
                    {
                        if (item.Name.Substring(0,3) == "IA_")
                        {
                            tag = item.RefersToRange.Cells.get_Address();

                            if (tag == objRange.Address)
                            {

                                if ((NroFilaPrincipal - 1) > 0)
                                {
                                    objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal - 1, 1];
                                    iTotalColumns = ActiveWorksheet.UsedRange.Columns.Count;

                                    while (k <= iTotalColumns)
                                    {

                                        if (objRange.Cells[k].HasFormula)
                                            tieneformula = true;

                                        k = k + 1;
                                    }
                                }
                                puedeinsertar = true;


                                break;
                            }
                        }
                    }
                    if (puedeinsertar)
                    {
                        Indices NewIndices = new Indices(NroFilaPrincipal, tieneformula);
                        NewIndices.ShowDialog();

                    }
                    else
                    {
                        objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal, 2];
                        var ConceptoPrevio = objRange.get_Value(Type.Missing);
                        List<oConcepto> ConceptVal = new List<oConcepto>();
                        ConceptVal = Generales.DameConceptosValidos();
                        bool CncValido = false;
                        if (ConceptoPrevio != null)
                        {
                            ConceptoPrevio = ConceptoPrevio.ToString();


                            CncValido = Generales.EsConceptoValido(ConceptoPrevio);

                            if (CncValido)  //(ConceptoPrevio.Substring(0, 4).ToUpper() == "OTRO") 
                            {
                                NroFilaPrincipal = objRange.Row;
                                NroColPrincipal = objRange.Column;
                                if ((NroFilaPrincipal - 1) > 0)
                                {

                                    var RangeConFr = ActiveWorksheet.get_Range(string.Format("{0}:{0}", NroFilaPrincipal - 1, Type.Missing));
                                    iTotalColumns = ActiveWorksheet.UsedRange.Columns.Count;

                                    while (k <= iTotalColumns)
                                    {

                                        if (RangeConFr.Cells[k].HasFormula)
                                            tieneformula = true;

                                        k = k + 1;
                                    }
                                }

                                Indices NewIndices = new Indices(NroFilaPrincipal, tieneformula);
                                NewIndices.Show();
                            }
                            else
                                MessageBox.Show("No es posible agregar índices debajo del índice " + IndicePrevio, "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        }


                    }
                }
                else
                {
                    MessageBox.Show("No es posible agregar índices debajo del índice EXPLICACION", "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("No es posible agregar índices en la fila seleccionada", "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                

            }
        }

        private void btnPrellenar_Click(object sender, RibbonControlEventArgs e)
        {
            string _Path = ExcelAddIn.Access.Configuration.Path;

            var fecha = DateTime.Now;
            var name = "Cruce_" + fecha.Year.ToString() + fecha.Month.ToString() + fecha.Day.ToString() + fecha.Hour.ToString() + fecha.Minute.ToString() + fecha.Second.ToString();
            var filepath = _Path + "\\" + name + ".pdf";
            // Creamos el documento con el tamaño de página tradicional
            Document doc = new Document(PageSize.LETTER);
            // Indicamos donde vamos a guardar el documento
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(filepath, FileMode.Create));
            // Le colocamos el título y el autor
            doc.AddTitle("Cruces");
            doc.AddCreator("S-DAT");
            // Abrimos el archivo
            doc.Open();
            PdfPTable tabla = new PdfPTable(3);
            for (int i = 0; i < 15; i++)
            {
                
                tabla.AddCell("A " + i);
                tabla.AddCell("B " + i);
            }
            doc.Add(tabla);

            doc.Close();
            writer.Close();
        }

        private void btnEliminarIndice_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sheetControl = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            string IndiceActivo = "";
            string IndiceSiguiente = "";
            bool Eliminar = false;
            List<string> NombreRangos = new List<string>();
            List<string> NombreRangosDEL = new List<string>();
            List<int> FilaPadre = new List<int>();
            int FilapadreAux = 0;
            long dif = 0;
            string NamedRange = "";
            bool tienedif = false;
            Excel.Range objRange = null;
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.Selection; // filas seleccionadas


            try
            {

                foreach (Excel.Range cell in currentCell.Cells)
                {

                    try
                    {


                        foreach (Excel.Name item1 in wb.Names)
                        {
                            // comparo la direccion de la celda con la del nombre del rango
                            if (item1.Name.Substring(0, 3) == "IA_")
                            {
                                if (item1.RefersToRange.Cells.get_Address() == cell.Address)
                                {
                                    NamedRange = item1.Name;

                                    break;
                                }
                            }
                        }

                        FilapadreAux = cell.Row;

                        if (!FilaPadre.Contains(FilapadreAux))
                            FilaPadre.Add(FilapadreAux);


                        objRange = (Excel.Range)sheet.Cells[cell.Row, 1];
                        IndiceActivo = objRange.Value2;


                        if (IndiceActivo.ToUpper().Trim() == "EXPLICACION")
                        {
                            MessageBox.Show("No es posible eliminar el índice EXPLICACION.", "Eliminar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Eliminar = false;
                            break;
                        }
                        if ((NamedRange != "IA_" + IndiceActivo) || (NamedRange == ""))
                        {
                            MessageBox.Show("No es posible eliminar un índice de formato guía", "Eliminar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            Eliminar = false;
                            break;
                        }
                        else
                        {
                            Eliminar = true;

                            NombreRangosDEL.Add("IA_" + IndiceActivo);
                        }


                    }
                    catch (Exception ex)
                    {

                        //MessageBox.Show(ex.Message);

                    }

                }

                if (Eliminar)
                {


                    currentCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;


                    objRange = (Excel.Range)sheet.Cells[currentCell.Cells.Row + 1, 1];
                    IndiceSiguiente = objRange.Value2;
                    if (IndiceSiguiente.ToUpper().Trim() == "EXPLICACION")
                        objRange.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    currentCell.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    NombreRangosDEL.Sort();
                    string NM = NombreRangosDEL.FirstOrDefault();
                    sheetControl.Controls.Remove(NM);
                    foreach (Excel.Name item2 in wb.Names)
                        NombreRangos.Add(item2.Name);
                    string[] split = NM.Split('_');
                    NM = split[1];
                    // foreach (string Nm in NombreRangosDEL)
                    long NamedRng = Convert.ToInt64(NM) + 100;
                    string IndiceSig = "0" + Convert.ToString(NamedRng);
                    while (NombreRangos.Contains("IA_" + IndiceSig))
                    {
                        sheetControl.Controls.Remove("IA_" + IndiceSig);

                        NamedRng = Convert.ToInt64(IndiceSig) + 100;
                        IndiceSig = "0" + Convert.ToString(NamedRng);
                    }
                    FilaPadre.Sort();
                    int row = FilaPadre.FirstOrDefault();

                    objRange = (Excel.Range)sheet.Cells[row, 1];
                    IndiceActivo = objRange.get_Value(Type.Missing).ToString();

                    objRange = (Excel.Range)sheet.Cells[row - 1, 1];
                    string IndiceAnt = objRange.get_Value(Type.Missing).ToString();
                    //i = 1;

                    while (NombreRangos.Contains("IA_" + IndiceActivo))
                    {
                        tienedif = false;

                        dif = Convert.ToInt64(IndiceActivo) - Convert.ToInt64(IndiceAnt);
                        while (dif != 100)
                        {

                            IndiceAnt = "0" + Convert.ToString(Convert.ToInt64(IndiceActivo) - 100);
                            IndiceActivo = IndiceAnt;

                            dif = dif - 100;

                            tienedif = true;
                        }

                        objRange = (Excel.Range)sheet.Cells[row, 1];
                        objRange.Value2 = IndiceAnt;

                        if (tienedif)
                            Generales.AddNamedRange(row, 1, "IA_" + Convert.ToString(IndiceAnt));
                        //busco el siguiente activo
                        row++;
                        objRange = (Excel.Range)sheet.Cells[row, 1];
                        IndiceActivo = objRange.get_Value(Type.Missing).ToString();


                    }


                    row = Generales.DameRangoPrincipal(FilaPadre.FirstOrDefault(), sheet);// busco el numero de fila OTRO para agregarle luego la sumatoria de los indices nuevos


                    Excel.Range objRangeJ = ((Excel.Range)sheet.Cells[FilaPadre[0], 1]);
                    objRangeJ.Select();
                    try
                    { // limpio si hay error en la formula
                        Excel.Range objRangeI = ((Excel.Range)sheet.Cells[row, 1]).SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors);//obten las celdas con errores

                        //////select all the cells with error formula
                        objRangeI.Clear();
                        objRangeI.Select();
                    }
                    catch (Exception ex)
                    {

                    }



                }



            }
            catch (Exception ex)
            {

                //  MessageBox.Show(ex.Message);

            }



        }
    }
}