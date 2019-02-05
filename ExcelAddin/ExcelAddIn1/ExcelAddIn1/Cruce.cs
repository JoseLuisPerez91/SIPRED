using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace ExcelAddIn1 {
    public partial class Cruce : Base {
        public Cruce() {
            InitializeComponent();
        }

        private void btnAceptar_Click(object sender, EventArgs e) {
            string _Path = ExcelAddIn.Access.Configuration.Path;
            oTipoPlantilla[] _TemplateTypes = Assembler.LoadJson<oTipoPlantilla[]>($"{_Path}\\jsons\\TiposPlantillas.json");
            oCruce[] _Cruces = Assembler.LoadJson<oCruce[]>($"{_Path}\\jsons\\Cruces.json");
            List<oCruce> _result = new List<oCruce>();
            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
            //FileInfo _Excel = new FileInfo($"{_Path}\\jsons\\SIPRED-EstadosFinancierosGeneral.xlsm");
            oTipoPlantilla _TemplateType = null;
            using (ExcelPackage _package = new ExcelPackage(_Excel))
            {
                foreach (oTipoPlantilla _TT in _TemplateTypes)
                {
                    if (_package.Workbook.Worksheets.Where(o => o.Name == _TT.Clave).FirstOrDefault() != null)
                        _TemplateType = _TT;
                }
                if (_TemplateType != null)
                {
                    ExcelWorksheet _wsTest = _package.Workbook.Worksheets.First(o => o.Name == _TemplateType.Clave);

                    foreach (oCruce _Cruce in _Cruces.Where(o => o.IdTipoPlantilla == _TemplateType.IdTipoPlantilla))
                    {
                        _Cruce.setCeldas();
                        foreach (oCelda _Celda in _Cruce.CeldasFormula)
                        {
                            ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Celda.Anexo];
                            int _maxValue = _workSheet.Dimension.Rows + 1;
                            int _maxRow = (_workSheet.Dimension.Rows / 2) + (_workSheet.Dimension.Rows % 2);
                            for (int i = 1; i <= _maxRow; i++)
                            {
                                _Celda.Fila = (_workSheet.Cells[i, 1].Text == _Celda.Indice) ? i : _Celda.Fila;
                                _Celda.Fila = (_workSheet.Cells[(_maxValue - i), 1].Text == _Celda.Indice) ? _maxValue - i : _Celda.Fila;
                                if (_Celda.Fila > -1)
                                {
                                    _Celda.setCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna], _Celda.Anexo);
                                    var celdas = _workSheet.Cells[_Celda.Fila, _Celda.Columna];
                                    _Celda.Concepto = _workSheet.Cells[_Celda.Fila, 2].Text;
                                    double valor1;
                                    if (double.TryParse(_workSheet.Cells[_Celda.Fila, _Celda.Columna].Text, out valor1))
                                        _Celda.Valor = valor1;
                                    else
                                    {
                                        var wscell = _workSheet.Cells[_Celda.CeldaXlsx];
                                        wscell.Calculate();
                                        if (double.TryParse(wscell.Text, out valor1))
                                            _Celda.Valor = valor1;
                                    }
                                }
                            }
                        }
                        foreach (oCeldaCondicion _Celda in _Cruce.CeldasCondicion)
                        {
                            ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Celda.Anexo];
                            if (_workSheet != null)
                            {
                                int _maxValue = _workSheet.Dimension.Rows + 1;
                                int _maxRow = (_workSheet.Dimension.Rows / 2) + (_workSheet.Dimension.Rows % 2);
                                for (int i = 1; i <= _maxRow; i++)
                                {
                                    _Celda.Fila = (_workSheet.Cells[i, 1].Text == _Celda.Indice) ? i : _Celda.Fila;
                                    _Celda.Fila = (_workSheet.Cells[(_maxValue - i), 1].Text == _Celda.Indice) ? _maxValue - i : _Celda.Fila;
                                    if (_Celda.Fila > -1)
                                    {
                                        _Celda.setCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna], _Celda.Anexo);
                                        _Celda.Concepto = _workSheet.Cells[_Celda.Fila, 2].Text;
                                        //var wscell = _workSheet.Cells[_Celda.CeldaExcel];
                                        //wscell.Calculate();
                                        //_Celda.Valor = wscell.Text;
                                    }
                                }
                            }
                        }
                       
                        _Cruce.setFormulaExcel();
                        var validar = false;
                        if (_Cruce.CondicionExcel != "")
                        {
                            var celda = _wsTest.Cells["A2"];
                            _wsTest.Cells["A2"].Formula = _Cruce.CondicionExcel;
                            _wsTest.Cells["A2"].Calculate();
                            _Cruce.ResultadoCondicion = _wsTest.Cells["A2"].Value.ToString();

                            if (_Cruce.ResultadoCondicion == "SI")
                                validar = true;
                            else
                                validar = false;
                        }
                        else { validar = true; }

                        if (validar)
                        {
                            bool result = false;
                            if (_Cruce.FormulaNumero.Contains(":"))
                            {
                                _Cruce.FormulaNumero = _Cruce.FormulaNumero.Replace(":", "+");
                                _Cruce.FormulaNumero = _Cruce.FormulaNumero.Replace("SUM", "");
                            }
                            var value1 = _Cruce.FormulaNumero.Split('=')[0];
                            var value2 = _Cruce.FormulaNumero.Split('=')[1];

                            var result1 = new DataTable().Compute(value1, null);
                            var result2 = new DataTable().Compute(value2, null);

                            _Cruce.Grupo1 = Convert.ToDouble(result1);
                            _Cruce.Grupo2 = Convert.ToDouble(result2);
                            _Cruce.Diferencia = Math.Abs(_Cruce.Grupo1 - _Cruce.Grupo2);

                            if (_Cruce.Diferencia > 0)
                                result = false;
                            else
                                result = true;

                            if (!result)
                                _result.Add(_Cruce);

                        }
                    }
                }
                else if (_TemplateType == null)
                    MessageBox.Show("Archivo no valido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (_result.Count > 0)
                CreatePDF(_result.ToArray(), _Cruces, _Path);
            //    ReadFormula(_result.ToArray(), _Excel.FullName, index, _Cruces);
            else
                MessageBox.Show("No se encontraron diferencias", "Información Correcta", MessageBoxButtons.OK, MessageBoxIcon.Information);


            this.Hide();
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Hide();
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
            doc.AddTitle("Curces");
            doc.AddCreator("D-SAT");
            // Abrimos el archivo
            doc.Open();
            // Creamos el tipo de Font que vamos utilizar
            iTextSharp.text.Font titlefont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 8, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            iTextSharp.text.Font _standardFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font _standardFontbold = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 7, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            // Escribimos el encabezado en el documento
            doc.Add(new Paragraph("eISSIF XML 17"));
            doc.Add(new Paragraph("Cruces", _standardFont));
            doc.Add(new Paragraph("SIPRED - ESTADOS FINANCIEROS GENERAL"));

            PdfPTable tblHeader = new PdfPTable(7);
            tblHeader.WidthPercentage = 100;
            PdfPCell cellNum = new PdfPCell(new Phrase("Número", titlefont));
            cellNum.BorderWidth = 0;
            cellNum.BorderWidthTop = 0.75f;
            cellNum.BorderWidthBottom = 0.75f;
            cellNum.BorderColorTop = new BaseColor(Color.Blue);
            cellNum.BorderColorBottom = new BaseColor(Color.White);

            PdfPCell cellconc = new PdfPCell(new Phrase("Concepto", titlefont));
            cellconc.BorderWidth = 0;
            cellconc.BorderWidthTop = 0.75f;
            cellconc.BorderWidthBottom = 0.75f;
            cellconc.BorderColorTop = new BaseColor(Color.Blue);
            cellconc.BorderColorBottom = new BaseColor(Color.White);
            cellconc.Colspan = 6;

            tblHeader.AddCell(cellNum);
            tblHeader.AddCell(cellconc);

            PdfPCell col1 = new PdfPCell(new Phrase("", titlefont));
            col1.BorderWidth = 0;
            col1.BorderWidthTop = 0.75f;
            col1.BorderWidthBottom = 0.75f;
            col1.BorderColorBottom = new BaseColor(Color.Blue);
            col1.BorderColorTop = new BaseColor(Color.White);

            PdfPCell col2 = new PdfPCell(new Phrase("Índice", titlefont));
            col2.BorderWidth = 0;
            col2.BorderWidthTop = 0.75f;
            col2.BorderWidthBottom = 0.75f;
            col2.BorderColorBottom = new BaseColor(Color.Blue);
            col2.BorderColorTop = new BaseColor(Color.White);

            PdfPCell col3 = new PdfPCell(new Phrase("Col.", titlefont));
            col3.BorderWidth = 0;
            col3.BorderWidthTop = 0.75f;
            col3.BorderWidthBottom = 0.75f;
            col3.BorderColorBottom = new BaseColor(Color.Blue);
            col3.BorderColorTop = new BaseColor(Color.White);

            PdfPCell col4 = new PdfPCell(new Phrase("Concepto", titlefont));
            col4.BorderWidth = 0;
            col4.BorderWidthTop = 0.75f;
            col4.BorderWidthBottom = 0.75f;
            col4.BorderColorBottom = new BaseColor(Color.Blue);
            col4.BorderColorTop = new BaseColor(Color.White);
            col4.Colspan = 2;

            PdfPCell col6 = new PdfPCell(new Phrase("Gpo. 1", titlefont));
            col6.BorderWidth = 0;
            col6.BorderWidthTop = 0.75f;
            col6.BorderWidthBottom = 0.75f;
            col6.BorderColorBottom = new BaseColor(Color.Blue);
            col6.BorderColorTop = new BaseColor(Color.White);

            PdfPCell col7 = new PdfPCell(new Phrase("Gpo. 2", titlefont));
            col7.BorderWidth = 0;
            col7.BorderWidthTop = 0.75f;
            col7.BorderWidthBottom = 0.75f;
            col7.BorderColorBottom = new BaseColor(Color.Blue);
            col7.BorderColorTop = new BaseColor(Color.White);

            tblHeader.AddCell(col1);
            tblHeader.AddCell(col2);
            tblHeader.AddCell(col3);
            tblHeader.AddCell(col4);
            tblHeader.AddCell(col6);
            tblHeader.AddCell(col7);
            doc.Add(Chunk.NEWLINE);
            foreach (var item in _result)
            {
                PdfPCell cellid = new PdfPCell(new Phrase(item.IdCruce.ToString(), titlefont));
                cellid.BorderWidth = 0;
                cellid.BorderWidthTop = 1;
                cellid.BorderColorTop = new BaseColor(Color.White);
                cellid.BackgroundColor = new BaseColor(Color.Gray);

                var strConcepto = cruces.Where(c => c.IdCruce == item.IdCruce).FirstOrDefault();
                PdfPCell cellconcepto = new PdfPCell(new Phrase(strConcepto.Concepto, titlefont));
                cellconcepto.BorderWidth = 0;
                cellconcepto.BorderWidthTop = 1;
                cellconcepto.BorderColorTop = new BaseColor(Color.White);
                cellconcepto.Colspan = 6;
                cellconcepto.BackgroundColor = new BaseColor(Color.Gray);

                tblHeader.AddCell(cellid);
                tblHeader.AddCell(cellconcepto);

                PdfPCell cellformula = new PdfPCell(new Phrase(item.Formula, _standardFont));
                cellformula.BorderWidth = 0;
                cellformula.Colspan = 7;
                tblHeader.AddCell(cellformula);

                if (item.Condicion != null || item.Condicion.Length > 0)
                {
                    PdfPCell cellcondicion = new PdfPCell(new Phrase(item.Condicion, _standardFont));
                    cellcondicion.BorderWidth = 0;
                    cellcondicion.Colspan = 7;
                    tblHeader.AddCell(cellcondicion);
                }

                var formula1 = item.Formula.Split('=')[0];
                var formula2 = item.Formula.Split('=')[1];

                var valor = 1;
                foreach (var detail in item.CeldasFormula)
                {
                    var color = Color.LightGray;

                    if ((valor % 2) == 0)
                        color = Color.White;

                    PdfPCell cellanexo = new PdfPCell(new Phrase(detail.Anexo, _standardFont));
                    cellanexo.BorderWidth = 0;
                    cellanexo.BackgroundColor = new BaseColor(color);
                    PdfPCell cellindice = new PdfPCell(new Phrase(detail.Indice, _standardFont));
                    cellindice.BorderWidth = 0;
                    cellindice.BackgroundColor = new BaseColor(color);
                    PdfPCell cellcolumna = new PdfPCell(new Phrase(detail.Columna.ToString(), _standardFont));
                    cellcolumna.BorderWidth = 0;
                    cellcolumna.BackgroundColor = new BaseColor(color);
                    PdfPCell cellconceptodet = new PdfPCell(new Phrase(detail.Concepto, _standardFont));
                    cellconceptodet.BorderWidth = 0;
                    cellconceptodet.BackgroundColor = new BaseColor(color);
                    cellconceptodet.Colspan = 2;

                    var strgpo1 = string.Empty;
                    var strgpo2 = string.Empty;

                    if (formula1.Contains(detail.Original))
                        strgpo1 = detail.Valor == 0 ? "" : detail.Valor.ToString("C");

                    if (formula2.Contains(detail.Original))
                        strgpo2 = detail.Valor == 0 ? "" : detail.Valor.ToString("C");

                    PdfPCell cellgpo1 = new PdfPCell(new Phrase(strgpo1, _standardFont));
                    cellgpo1.BorderWidth = 0;
                    cellgpo1.BackgroundColor = new BaseColor(color);
                    cellgpo1.HorizontalAlignment = Element.ALIGN_RIGHT;

                    PdfPCell cellgpo2 = new PdfPCell(new Phrase(strgpo2, _standardFont));
                    cellgpo2.BorderWidth = 0;
                    cellgpo2.BackgroundColor = new BaseColor(color);
                    cellgpo2.HorizontalAlignment = Element.ALIGN_RIGHT;

                    tblHeader.AddCell(cellanexo);
                    tblHeader.AddCell(cellindice);
                    tblHeader.AddCell(cellcolumna);
                    tblHeader.AddCell(cellconceptodet);
                    tblHeader.AddCell(cellgpo1);
                    tblHeader.AddCell(cellgpo2);

                    valor++;
                }

                foreach(var gen in item.CeldasCondicion)
                {
                    var color = Color.LightGray;

                    if ((valor % 2) == 0)
                        color = Color.White;

                    PdfPCell cellgs = new PdfPCell(new Phrase(gen.Anexo, _standardFont));
                    cellgs.BorderWidth = 0;
                    cellgs.BackgroundColor = new BaseColor(color);

                    PdfPCell cellgi = new PdfPCell(new Phrase(gen.Indice, _standardFont));
                    cellgi.BorderWidth = 0;
                    cellgi.BackgroundColor = new BaseColor(color);

                    PdfPCell cellgcol = new PdfPCell(new Phrase(gen.Columna.ToString(), _standardFont));
                    cellgcol.BorderWidth = 0;
                    cellgcol.BackgroundColor = new BaseColor(color);

                    PdfPCell cellgconc = new PdfPCell(new Phrase(gen.Concepto == null ? " " : gen.Concepto, _standardFont));
                    cellgconc.BorderWidth = 0;
                    cellgconc.BackgroundColor = new BaseColor(color);
                    cellconc.Colspan = 3;

                    PdfPCell cellgv = new PdfPCell(new Phrase(gen.Valor == null ? " " : gen.Valor, _standardFont));
                    cellgv.BorderWidth = 0;
                    cellgv.BackgroundColor = new BaseColor(color);

                    tblHeader.AddCell(cellgs);
                    tblHeader.AddCell(cellgi);
                    tblHeader.AddCell(cellgcol);
                    tblHeader.AddCell(cellgconc);
                    tblHeader.AddCell(cellgv);

                    valor++;
                }

                PdfPCell cellcalc = new PdfPCell(new Phrase("Cálculos", _standardFontbold));
                cellcalc.BorderWidth = 0;
                cellcalc.HorizontalAlignment = Element.ALIGN_RIGHT;
                cellcalc.Colspan = 5;

                PdfPCell cellgpot1 = new PdfPCell(new Phrase(item.Grupo1.ToString("C"), _standardFont));
                cellgpot1.BorderWidth = 0;
                cellgpot1.HorizontalAlignment = Element.ALIGN_RIGHT;

                PdfPCell cellgpot2 = new PdfPCell(new Phrase(item.Grupo2.ToString("C"), _standardFont));
                cellgpot2.BorderWidth = 0;
                cellgpot2.HorizontalAlignment = Element.ALIGN_RIGHT;

                tblHeader.AddCell(cellcalc);
                tblHeader.AddCell(cellgpot1);
                tblHeader.AddCell(cellgpot2);

                PdfPCell celldifempty = new PdfPCell(new Phrase(" ", _standardFont));
                celldifempty.BorderWidth = 1;
                celldifempty.BorderColor = new BaseColor(Color.White);
                celldifempty.Colspan = 5;

                PdfPCell celldifText = new PdfPCell(new Phrase("Diferencia", _standardFontbold));
                celldifText.BorderWidth = 1;
                celldifText.BorderColor = new BaseColor(Color.White);
                celldifText.BackgroundColor = new BaseColor(Color.LightGray);

                PdfPCell celldif = new PdfPCell(new Phrase(item.Diferencia.ToString("C"), _standardFontbold));
                celldif.BorderWidth = 1;
                celldif.BorderColor = new BaseColor(Color.White);
                celldif.HorizontalAlignment = Element.ALIGN_RIGHT;
                celldifText.BackgroundColor = new BaseColor(Color.LightGray);

                tblHeader.AddCell(celldifempty);
                tblHeader.AddCell(celldifText);
                tblHeader.AddCell(celldif);
            }

            doc.Add(tblHeader);

            doc.Close();
            writer.Close();

            WebBrowser wb = new WebBrowser();
            wb.Navigate(filepath);
        }

        private bool ValidateOperation(string formula)
        {
            var result = false;
            double result1;
            double result2;
            double total1;
            double total2;
            var operacion1 = formula.Split('=')[0];
            var operacion2 = formula.Split('=')[1];

            if(double.TryParse(operacion1, out result1))
            {
                total1 = result1;
            }
            else
            {

            }

           if(!double.TryParse(operacion2, out result2))
            {
                total2 = result2;
            }
            else
            {

            }

            return result;
        }
    }
}