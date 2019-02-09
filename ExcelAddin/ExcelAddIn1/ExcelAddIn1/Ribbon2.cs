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
                        if (item.Name.Substring(0,3) == "IA_0")
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
    }
}