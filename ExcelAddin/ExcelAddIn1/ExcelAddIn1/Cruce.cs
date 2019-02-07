using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class Cruce : Base
    {
        public Cruce()
        {
            InitializeComponent();
        }

        private void btnAceptar_Click(object sender, EventArgs e)
        {
            //this.Hide();

            string _Path = ExcelAddIn.Access.Configuration.Path;
            bool _File = false;

            if (Directory.Exists(_Path + "\\jsons") && Directory.Exists(_Path + "\\templates"))
            {
                if (File.Exists(_Path + "\\jsons\\ValidacionCruces.json"))
                {
                    _File = true;
                }
                else
                {
                    MessageBox.Show("Los archivos base serán generados... Click en el botón Aceptar para continuar.", "Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    KeyValuePair<bool, string[]> _results = new lSerializados().ObtenerSerializados();
                    _File = true;
                }
            }
            else
            {
                MessageBox.Show("Los archivos base serán generados... Click en el botón Aceptar para continuar.", "Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (!Directory.Exists(_Path + "\\jsons"))
                {
                    Directory.CreateDirectory(_Path + "\\jsons");
                }
                if (!Directory.Exists(_Path + "\\templates"))
                {
                    Directory.CreateDirectory(_Path + "\\templates");
                }

                KeyValuePair<bool, string[]> _results = new lSerializados().ObtenerSerializados();
                _File = true;
            }

            oValidaCruces[] _ValidaCruces = null;
            if (_File)
            {
                _ValidaCruces = Assembler.LoadJson<oValidaCruces[]>($"{_Path}\\jsons\\ValidacionCruces.json");

            }

            if(_ValidaCruces == null)
            {
                MessageBox.Show("No existe el archivo base, Generé uno nuevo para crear los archivos base...", "Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.Close();
            }
            if (!ValidaCruces(_ValidaCruces))
            {
                this.Hide();
                return;
            }

            this.pgbCruces.Visible = true;
            lblTitle.Text = "Comienzo verificación, por favor espere!! ";
            this.btnAceptar.Visible = false;
            this.btnCancelar.Visible = false;
            int progress = 0;
            progress += 10;
            pgbCruces.Value = progress;
            oTipoPlantilla[] _TemplateTypes = Assembler.LoadJson<oTipoPlantilla[]>($"{_Path}\\jsons\\TiposPlantillas.json");
            oCruce[] _Cruces = Assembler.LoadJson<oCruce[]>($"{_Path}\\jsons\\Cruces.json");

            List<oCruce> _result = new List<oCruce>();
            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);

            progress += 10;
            pgbCruces.Value = progress;
            //FileInfo _Excel = new FileInfo($"{_Path}\\jsons\\SIPRED-EstadosFinancierosGeneral.xlsm");
            oTipoPlantilla _TemplateType = null;
            using (ExcelPackage _package = new ExcelPackage(_Excel))
            {
                foreach (oTipoPlantilla _TT in _TemplateTypes)
                {
                    if (_package.Workbook.Worksheets.Where(o => o.Name == _TT.Clave).FirstOrDefault() != null)
                        _TemplateType = _TT;
                }
                progress += 10;
                pgbCruces.Value = progress;
                if (_TemplateType != null)
                {
                    //_package.Workbook.Worksheets.Add("Test");
                    // ExcelWorksheet _wsTest = _package.Workbook.Worksheets.First(o => o.Name == "Test");
                    //INTEROOP//
                    Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                    Worksheet xlSht = null;
                    Range currentCell = null;
                    Range currentFind = null;
                    ////////
                    foreach (oCruce _Cruce in _Cruces.Where(o => o.IdTipoPlantilla == _TemplateType.IdTipoPlantilla))
                    {
                        _Cruce.setCeldas();
                        foreach (oCelda _Celda in _Cruce.CeldasFormula)
                        {
                            ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Celda.Anexo];
                            lblTitle.Text = "Verificando " + _Celda.Anexo;
                            if (_workSheet != null)
                            {
                                int _maxValue = _workSheet.Dimension.Rows + 1;
                                int _maxRow = (_workSheet.Dimension.Rows / 2) + (_workSheet.Dimension.Rows % 2);


                                xlSht = (Worksheet)wb.Worksheets.get_Item(_Celda.Anexo);

                                currentCell = (Range)xlSht.get_Range("A1", "A" + (_maxValue).ToString());


                                currentFind = currentCell.Find(_Celda.Indice, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                                   XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
                                    Type.Missing, Type.Missing);
                                if (currentFind != null)
                                {
                                    _Celda.Fila = currentFind.Row;

                                    _Celda.setFullAddressCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna]);
                                    _Celda.Concepto = _workSheet.Cells[_Celda.Fila, 2].Text;

                                    currentCell = (Range)xlSht.Cells[_Celda.Fila, _Celda.Columna];
                                    if (currentCell.get_Value(Type.Missing) != null)
                                        _Celda.Valor = currentCell.get_Value(Type.Missing).ToString();
                                    else
                                        _Celda.Valor = "0";


                                }
                            }


                        }

                        pgbCruces.Value = progress;
                        foreach (oCeldaCondicion _Celda in _Cruce.CeldasCondicion)
                        {
                            ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Celda.Anexo];
                            if (_workSheet != null)
                            {
                                int _maxValue = _workSheet.Dimension.Rows + 1;
                                int _maxRow = (_workSheet.Dimension.Rows / 2) + (_workSheet.Dimension.Rows % 2);

                                xlSht = (Worksheet)wb.Worksheets.get_Item(_Celda.Anexo);

                                currentCell = (Range)xlSht.get_Range("A1", "A" + (_maxValue).ToString());


                                currentFind = currentCell.Find(_Celda.Indice, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                                   XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
                                    Type.Missing, Type.Missing);
                                if (currentFind != null)
                                {
                                    _Celda.Fila = currentFind.Row;
                                    _Celda.setFullAddressCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna]);

                                }


                            }
                        }


                        _Cruce.setFormulaExcel();


                        xlSht = (Worksheet)wb.Worksheets.get_Item("SIPRED");
                        Range Test_Range = (Range)xlSht.get_Range("A1");
                        string ValorAnterior = Test_Range.get_Value(Type.Missing);
                        string[] formula;



                        Test_Range.Formula = "=" + _Cruce.FormulaExcel;

                        _Cruce.ResultadoFormula = Test_Range.get_Value(Type.Missing).ToString();

                        xlSht.Cells[1, 1] = ValorAnterior;// restauro




                        if (_Cruce.CondicionExcel != "")
                        {
                            Test_Range = (Range)xlSht.get_Range("A2");
                            ValorAnterior = Test_Range.get_Value(Type.Missing);
                            Test_Range.Formula = "=" + _Cruce.CondicionExcel;
                            _Cruce.ResultadoCondicion = Test_Range.get_Value(Type.Missing).ToString();
                            xlSht.Cells[2, 1] = ValorAnterior;// restauro



                            _Cruce.Condicion = "[" + _Cruce.Condicion + "] = " + _Cruce.ResultadoCondicion;
                        }



                        if (_Cruce.ResultadoFormula.ToLower() == "false")
                        {

                            //calculo la diferencia
                            if (_Cruce.FormulaExcel.Contains("="))
                            {
                                formula = _Cruce.FormulaExcel.Split('=');
                                Test_Range = (Range)xlSht.get_Range("A3");
                                ValorAnterior = Test_Range.get_Value(Type.Missing);
                                Test_Range.Formula = "=" + formula[0] + "-" + formula[1];
                                _Cruce.Diferencia = Test_Range.get_Value(Type.Missing).ToString();
                                xlSht.Cells[3, 1] = ValorAnterior;// restauro

                            }
                            _result.Add(_Cruce);
                        }

                        if (progress <= 70)
                        {
                            progress += 10;
                            pgbCruces.Value = progress;
                        }
                    }

                    progress += 5;

                    pgbCruces.Value = progress;
                }
                else if (_TemplateType == null)
                    MessageBox.Show("Archivo no válido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            progress += 15;

            pgbCruces.Value = progress;
            if (_result.Count > 0)
                CreatePDF(_result.ToArray(), _Cruces, _Path);
            else
                MessageBox.Show("No se encontraron diferencias", "Información Correcta", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Hide();
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        public static string ColumnAdress(int col)
        {
            if (col <= 26)
            {
                return Convert.ToChar(col + 64).ToString();
            }
            int div = col / 26;
            int mod = col % 26;
            if (mod == 0) { mod = 26; div--; }
            return ColumnAdress(div) + ColumnAdress(mod);
        }
        private void CreatePDF(oCruce[] _result, oCruce[] cruces, string path)
        {
            var fecha = DateTime.Now;
            var name = "Cruce_" + fecha.Year.ToString() + fecha.Month.ToString() + fecha.Day.ToString() + fecha.Hour.ToString() + fecha.Minute.ToString() + fecha.Second.ToString();
            var filepath = path + "\\" + name + ".pdf";
            // Creamos el documento con el tamaño de página tradicional
            Document doc = new Document(PageSize.LETTER);
            // Indicamos donde vamos a guardar el documento
            PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(filepath, FileMode.Create));
            // Le colocamos el título y el autor
            // **Nota: Esto no será visible en el documento
            doc.AddTitle("Cruces");
            doc.AddCreator("S-DAT");
            // Abrimos el archivo
            doc.Open();
            // Creamos el tipo de Font que vamos utilizar
            iTextSharp.text.Font titlefont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 9, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            iTextSharp.text.Font _standardFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            // Escribimos el encabezamiento en el documento
            doc.Add(new Paragraph("eISSIF XML 17"));
            doc.Add(new Paragraph("Cruces", _standardFont));
            doc.Add(new Paragraph("SIPRED - ESTADOS FINANCIEROS GENERAL"));
            doc.Add(Chunk.NEWLINE);

            PdfPTable tblHeader = new PdfPTable(5);
            tblHeader.WidthPercentage = 100;
            PdfPCell cellNum = new PdfPCell(new Phrase("Número", titlefont));
            cellNum.BorderWidth = 0;
            cellNum.BorderWidthTop = 0.75f;
            cellNum.BorderWidthBottom = 0.75f;
            cellNum.BorderColorTop = new BaseColor(Color.Blue);
            cellNum.BorderColorBottom = new BaseColor(Color.Blue);

            PdfPCell cellconc = new PdfPCell(new Phrase("Concepto", titlefont));
            cellconc.BorderWidth = 0;
            cellconc.BorderWidthTop = 0.75f;
            cellconc.BorderWidthBottom = 0.75f;
            cellconc.BorderColorTop = new BaseColor(Color.Blue);
            cellconc.BorderColorBottom = new BaseColor(Color.Blue);

            PdfPCell cellCol3 = new PdfPCell(new Phrase("", titlefont));
            cellCol3.BorderWidth = 0;
            cellCol3.BorderWidthTop = 0.75f;
            cellCol3.BorderWidthBottom = 0.75f;
            cellCol3.BorderColorTop = new BaseColor(Color.Blue);
            cellCol3.BorderColorBottom = new BaseColor(Color.Blue);

            PdfPCell cellCol4 = new PdfPCell(new Phrase("", titlefont));
            cellCol4.BorderWidth = 0;
            cellCol4.BorderWidthTop = 0.75f;
            cellCol4.BorderWidthBottom = 0.75f;
            cellCol4.BorderColorTop = new BaseColor(Color.Blue);
            cellCol4.BorderColorBottom = new BaseColor(Color.Blue);
            PdfPCell cellCol5 = new PdfPCell(new Phrase("", titlefont));
            cellCol5.BorderWidth = 0;
            cellCol5.BorderWidthTop = 0.75f;
            cellCol5.BorderWidthBottom = 0.75f;
            cellCol5.BorderColorTop = new BaseColor(Color.Blue);
            cellCol5.BorderColorBottom = new BaseColor(Color.Blue);

            tblHeader.AddCell(cellNum);
            tblHeader.AddCell(cellconc);
            tblHeader.AddCell(cellCol3);
            tblHeader.AddCell(cellCol4);
            tblHeader.AddCell(cellCol5);

            foreach (var item in _result)
            {




                PdfPCell cellid = new PdfPCell(new Phrase(item.IdCruce.ToString(), titlefont));
                cellid.BorderWidth = 0;

                var strConcepto = cruces.Where(c => c.IdCruce == item.IdCruce).FirstOrDefault();
                PdfPCell cellconcepto = new PdfPCell(new Phrase(strConcepto.Concepto + " (" + "DIFERENCIA: " + item.Diferencia + ")", titlefont));
                cellconcepto.BorderWidth = 0;
                cellconcepto.Colspan = 3;

                tblHeader.AddCell(cellid);
                tblHeader.AddCell(cellconcepto);

                PdfPCell cellformula = new PdfPCell(new Phrase(item.Formula, _standardFont));
                cellformula.BorderWidth = 0;
                cellformula.Colspan = 4;
                tblHeader.AddCell(cellformula);



                if (item.Condicion != null || item.Condicion.Length > 0)
                {
                    PdfPCell cellcondicion = new PdfPCell(new Phrase(item.Condicion, _standardFont));
                    cellcondicion.BorderWidth = 0;
                    cellcondicion.Colspan = 4;
                    tblHeader.AddCell(cellcondicion);
                }
                PdfPCell cellvalor = new PdfPCell(new Phrase("", _standardFont));
                cellvalor.BorderWidth = 0;
                cellformula.Colspan = 1;
                tblHeader.AddCell(cellvalor);


                PdfPCell cellanexohdr = new PdfPCell(new Phrase("Anexo", _standardFont));
                cellanexohdr.BorderWidth = 0;
                PdfPCell cellindicehdr = new PdfPCell(new Phrase("Indice", _standardFont));
                cellindicehdr.BorderWidth = 0;
                PdfPCell cellcolumnahdr = new PdfPCell(new Phrase("Columna", _standardFont));
                cellcolumnahdr.BorderWidth = 0;
                PdfPCell cellconceptodethdr = new PdfPCell(new Phrase("Concepto", _standardFont));
                cellconceptodethdr.BorderWidth = 0;
                PdfPCell cellGpo1thdr = new PdfPCell(new Phrase("Valor.", _standardFont));
                cellGpo1thdr.BorderWidth = 0;

                tblHeader.AddCell(cellanexohdr);
                tblHeader.AddCell(cellindicehdr);
                tblHeader.AddCell(cellcolumnahdr);
                tblHeader.AddCell(cellconceptodethdr);
                tblHeader.AddCell(cellGpo1thdr);

                var valor = 1;
                foreach (var detail in item.CeldasFormula)
                {
                    var color = Color.White;

                    if ((valor % 2) == 0)
                        color = Color.LightGray;

                    PdfPCell cellanexo = new PdfPCell(new Phrase(detail.Anexo, _standardFont));
                    cellanexo.BorderWidth = 0;
                    cellanexo.BackgroundColor = new BaseColor(color);
                    PdfPCell cellindice = new PdfPCell(new Phrase(detail.Indice, _standardFont));
                    cellindice.BorderWidth = 0;
                    cellindice.BackgroundColor = new BaseColor(color);
                    PdfPCell cellcolumna = new PdfPCell(new Phrase(ColumnAdress(detail.Columna), _standardFont));
                    cellcolumna.BorderWidth = 0;
                    cellcolumna.BackgroundColor = new BaseColor(color);
                    PdfPCell cellconceptodet = new PdfPCell(new Phrase(detail.Concepto, _standardFont));
                    cellconceptodet.BorderWidth = 0;
                    cellconceptodet.BackgroundColor = new BaseColor(color);

                    PdfPCell valorcolumna = new PdfPCell(new Phrase(detail.Valor, _standardFont));
                    valorcolumna.BorderWidth = 0;
                    valorcolumna.BackgroundColor = new BaseColor(color);


                    tblHeader.AddCell(cellanexo);
                    tblHeader.AddCell(cellindice);
                    tblHeader.AddCell(cellcolumna);
                    tblHeader.AddCell(cellconceptodet);
                    tblHeader.AddCell(valorcolumna);

                    valor++;
                }

            }

            doc.Add(tblHeader);

            doc.Close();
            writer.Close();

            WebBrowser wb = new WebBrowser();
            wb.Navigate(filepath);
        }

        public static bool ValidaCruces(oValidaCruces[] _ValidCruces)
        {
            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
            //INTEROOP//
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet xlSht = null;
            Range currentCell = null;
            Range currentFind = null;
            int _maxValue = 0;
            string _Hojas = "";
            ////////
            List<oValidaCruces> _Result = new List<oValidaCruces>();
            try
            {
                foreach (oValidaCruces _F in _ValidCruces)
                {
                    xlSht = (Worksheet)wb.Worksheets.get_Item(_F.Hoja);
                    if (xlSht != null)
                    {

                        _maxValue = xlSht.UsedRange.Count + 1;
                        currentCell = (Range)xlSht.get_Range("A1", "A" + (_maxValue).ToString());


                        currentFind = currentCell.Find(_F.Indice, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                           XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
                            Type.Missing, Type.Missing);
                        if (currentFind != null)
                        {
                            currentCell = (Range)xlSht.Cells[currentFind.Row, 3];
                            if (currentCell.get_Value(Type.Missing) == null)
                            {
                                _F.EsCorrecto = false;
                                currentCell = (Range)xlSht.Cells[currentFind.Row, 2];
                                _F.Concepto = currentCell.get_Value(Type.Missing);

                                if (!(_Hojas.Contains(_F.Hoja)))
                                    _Hojas = _F.Hoja + "," + _Hojas;

                                _Result.Add(_F);
                            }

                        }
                    }

                }
                if (_Result.Count() > 0)
                {

                    MessageBox.Show("Para que la verificación se realice correctamente es necesario completar o revisar respuestas en " + _Hojas.TrimEnd(',') + " de los índices relacionados a continuación", "Información Correcta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                return _Result.Count() == 0;
            }
            catch
            {
                MessageBox.Show("Archivo no válido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

        }
    }
}