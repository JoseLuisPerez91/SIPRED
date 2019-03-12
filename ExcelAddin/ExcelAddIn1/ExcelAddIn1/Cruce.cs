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
using ExcelAddIn.Access;
using iTextSharp;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace ExcelAddIn1
{
    public partial class Cruce : Base
    {
        int _TotalValidaciones;
        public Cruce()
        {
            string _Path = Configuration.Path;
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            InitializeComponent();

            if (Directory.Exists(_Path + "\\jsons") && Directory.Exists(_Path + "\\templates"))
            {
                if (File.Exists(_Path + "\\jsons\\TiposPlantillas.json"))
                {
                    if (_Connection)
                    {
                        KeyValuePair<bool, System.Data.DataTable> _TipoPlantilla = new lSerializados().ObtenerUpdate();

                        foreach (DataRow _Row in _TipoPlantilla.Value.Rows)
                        {
                            string _IdTipoPlantilla = _Row["IdTipoPlantilla"].ToString();
                            string _Fecha_Modificacion = _Row["Fecha_Modificacion"].ToString();
                            string _Linea = null;

                            if (File.Exists(_Path + "\\jsons\\Update" + _IdTipoPlantilla + ".txt"))
                            {
                                StreamReader sw = new StreamReader(_Path + "\\Jsons\\Update" + _IdTipoPlantilla + ".txt");
                                _Linea = sw.ReadLine();
                                sw.Close();

                                if (_Linea != null)
                                {
                                    if (_Linea != _Fecha_Modificacion)
                                    {
                                        this.TopMost = false;
                                        this.Enabled = false;
                                        this.Hide();
                                        FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                                        _FileJsonfrm._Form = this;
                                        _FileJsonfrm._Process = false;
                                        _FileJsonfrm._Update = true;
                                        _FileJsonfrm._window = this.Text;
                                        _FileJsonfrm.Show();
                                        return;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    if (!_Connection)
                    {
                        MessageBox.Show(_Message.Replace("...", ", para crear los archivos base..."), "Creación de Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.btnAceptar.Enabled = false;
                        return;
                    }
                    else
                    {
                        this.TopMost = false;
                        this.Enabled = false;
                        this.Hide();
                        FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                        _FileJsonfrm._Form = this;
                        _FileJsonfrm._Process = false;
                        _FileJsonfrm._Update = false;
                        _FileJsonfrm._window = this.Text;
                        _FileJsonfrm.Show();
                        return;
                    }
                }
            }
            else
            {
                if (!Directory.Exists(_Path + "\\jsons"))
                {
                    Directory.CreateDirectory(_Path + "\\jsons");
                }
                if (!Directory.Exists(_Path + "\\templates"))
                {
                    Directory.CreateDirectory(_Path + "\\templates");
                }

                this.TopMost = false;
                this.Enabled = false;
                this.Hide();
                FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                _FileJsonfrm._Form = this;
                _FileJsonfrm._Process = false;
                _FileJsonfrm._Update = false;
                _FileJsonfrm._window = this.Text;
                _FileJsonfrm.Show();
                return;
            }
        }
        public void btnAceptar_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn._result.Clear();
            string _Path = ExcelAddIn.Access.Configuration.Path;

            oValidaCruces[] _ValidaCruces = Assembler.LoadJson<oValidaCruces[]>($"{_Path}\\jsons\\ValidacionCruces.json");

            Generales.Proteccion(false);

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

            _TotalValidaciones = _Cruces.Count();

            //List<oCruce> _result = new List<oCruce>();
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
                    //INTEROOP//
                    Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                    Worksheet xlSht = null;
                    Range currentCell = null;
                    Range currentFind = null; Range oRng; Range vlrange;
                    string[] Formula;
                    string indice;
                    string DetalleFr;

                    try
                    {
                        foreach (oCruce _Cruce in _Cruces.Where(o => o.IdTipoPlantilla == _TemplateType.IdTipoPlantilla))
                        {
                            _Cruce.setCeldas();
                            List<oCelda> CeldaNws = new List<oCelda>();

                            foreach (oCelda _Celda in _Cruce.CeldasFormula)
                            {
                                ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Celda.Anexo];
                                lblTitle.Text = "Verificando " + _Celda.Anexo;
                                if (_workSheet != null)
                                {
                                    xlSht = (Worksheet)wb.Worksheets.get_Item(_Celda.Anexo);
                                    int _maxValue = xlSht.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
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
                                        {
                                            _Celda.Valor = currentCell.get_Value(Type.Missing).ToString();
                                        }
                                        else
                                        {
                                            _Celda.Valor = "0";
                                        }
                                    }

                                    int j = 0;
                                    if (_Cruce.Formula.Contains(":") && _Cruce.Formula.Contains("SUM"))
                                    {
                                        if (_Cruce.Formula.Contains("="))
                                        {
                                            Formula = _Cruce.Formula.Split('=');
                                            for (j = 0; j < Formula.Count(); j++)
                                            {
                                                if (Formula[j].Contains(":"))
                                                {
                                                    DetalleFr = Formula[j];
                                                    break;
                                                }
                                            }
                                        }

                                        if (_Cruce.CeldasFormula[j].Anexo == _Cruce.CeldasFormula[j + 1].Anexo)
                                        {
                                            ExcelWorksheet _workSheetAnx = _package.Workbook.Worksheets[_Cruce.CeldasFormula[j].Anexo];

                                            if (CeldaNws.Count() == 0)
                                            {
                                                CeldaNws = _Cruce.CeldasFormula.ToList();
                                            }

                                            if ((_Cruce.CeldasFormula[j].Fila != _Cruce.CeldasFormula[j + 1].Fila) && (_Cruce.CeldasFormula[j].Columna == _Cruce.CeldasFormula[j + 1].Columna))
                                            {
                                                for (int r = _Cruce.CeldasFormula[j].Fila; r < _Cruce.CeldasFormula[j + 1].Fila; r++)
                                                {
                                                    oRng = (Range)xlSht.Cells[r, 1];
                                                    if (oRng.get_Value(Type.Missing) != null)
                                                    {
                                                        indice = oRng.get_Value(Type.Missing).ToString();

                                                        if (_Cruce.CeldasFormula[j].Indice != indice && _Cruce.CeldasFormula[j + 1].Indice != indice)
                                                        {
                                                            oCelda CeldaNw = new oCelda();
                                                            CeldaNw.Fila = oRng.Row;
                                                            CeldaNw.Indice = indice;
                                                            CeldaNw.Columna = _Celda.Columna;
                                                            CeldaNw.Anexo = _Cruce.CeldasFormula[j].Anexo;
                                                            CeldaNw.Original = "";
                                                            CeldaNw.Grupo = j;
                                                            CeldaNw.Concepto = _workSheetAnx.Cells[CeldaNw.Fila, 2].Text;
                                                            vlrange = (Range)xlSht.Cells[r, CeldaNw.Columna];

                                                            if (vlrange.get_Value(Type.Missing) != null)
                                                            {
                                                                CeldaNw.Valor = vlrange.get_Value(Type.Missing).ToString();
                                                            }
                                                            else
                                                            {
                                                                CeldaNw.Valor = "0";
                                                            }

                                                            CeldaNw.setFullAddressCeldaExcel(_workSheetAnx.Cells[CeldaNw.Fila, CeldaNw.Columna]);
                                                            if (!CeldaNws.Contains(CeldaNw))
                                                            {
                                                                CeldaNws.Add(CeldaNw);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if ((_Cruce.CeldasFormula[j].Fila == _Cruce.CeldasFormula[j + 1].Fila && _Cruce.CeldasFormula[j].Fila != -1) && (_Cruce.CeldasFormula[j].Columna != _Cruce.CeldasFormula[j + 1].Columna))
                                                {
                                                    Worksheet xlShtAnx = (Worksheet)wb.Worksheets.get_Item(_Cruce.CeldasFormula[j].Anexo);
                                                    CeldaNws = new List<oCelda>();
                                                    if (CeldaNws.Count() == 0)
                                                    {
                                                        CeldaNws = _Cruce.CeldasFormula.ToList();
                                                    }

                                                    for (int r = _Cruce.CeldasFormula[j].Columna; r < _Cruce.CeldasFormula[j + 1].Columna; r++)
                                                    {

                                                        if (_Cruce.CeldasFormula[j].Columna != r && _Cruce.CeldasFormula[j + 1].Columna != r)
                                                        {
                                                            oCelda CeldaNw = new oCelda();
                                                            CeldaNw.Fila = _Cruce.CeldasFormula[j].Fila;
                                                            CeldaNw.Indice = _Cruce.CeldasFormula[j].Indice;
                                                            CeldaNw.Columna = r;
                                                            CeldaNw.Anexo = _Cruce.CeldasFormula[j].Anexo;
                                                            CeldaNw.Original = "";
                                                            CeldaNw.Grupo = j;
                                                            CeldaNw.Concepto = _Cruce.CeldasFormula[j].Concepto;
                                                            vlrange = (Range)xlShtAnx.Cells[_Cruce.CeldasFormula[j].Fila, r];

                                                            if (vlrange.get_Value(Type.Missing) != null)
                                                            {
                                                                CeldaNw.Valor = vlrange.get_Value(Type.Missing).ToString();
                                                            }
                                                            else
                                                            {
                                                                CeldaNw.Valor = "0";
                                                            }

                                                            CeldaNw.setFullAddressCeldaExcel(_workSheetAnx.Cells[CeldaNw.Fila, CeldaNw.Columna]);
                                                            if (!CeldaNws.Contains(CeldaNw))
                                                            {
                                                                CeldaNws.Add(CeldaNw);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                    }
                                }
                            }

                            pgbCruces.Value = progress;
                            foreach (oCeldaCondicion _Celda in _Cruce.CeldasCondicion)
                            {
                                ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Celda.Anexo];
                                if (_workSheet != null)
                                {
                                    xlSht = (Worksheet)wb.Worksheets.get_Item(_Celda.Anexo);
                                    int _maxValue = xlSht.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                                    currentCell = (Range)xlSht.get_Range("A1", "A" + (_maxValue).ToString());


                                    currentFind = currentCell.Find(_Celda.Indice, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                                       XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
                                        Type.Missing, Type.Missing);

                                    if (currentFind != null)
                                    {
                                        _Celda.Fila = currentFind.Row;
                                        _Celda.setFullAddressCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna]);
                                        _Celda.Concepto = _workSheet.Cells[_Celda.Fila, 2].Text;
                                    }
                                }
                            }
                            //catch
                            _Cruce.setFormulaExcel();

                            if (CeldaNws.Count() > 0)
                            {
                                _Cruce.CeldasFormula = CeldaNws.OrderBy(x => x.Indice).ToArray();
                            }

                            xlSht = (Worksheet)wb.Worksheets.get_Item("SIPRED");
                            Range Test_Range = (Range)xlSht.get_Range("A1");
                            string ValorAnterior = Test_Range.get_Value(Type.Missing);
                            string[] formula;
                            Test_Range.Formula = "=" + _Cruce.FormulaExcel;
                            _Cruce.ResultadoFormula = Test_Range.get_Value(Type.Missing).ToString();
                            xlSht.Cells[1, 1] = ValorAnterior;// restauro

                            if (_Cruce.FormulaExcel.Contains("="))
                            {
                                formula = _Cruce.FormulaExcel.Split('=');
                                Test_Range = (Range)xlSht.get_Range("A3");
                                ValorAnterior = Test_Range.get_Value(Type.Missing);
                                if (chksigno.Checked)
                                    Test_Range.Formula = "=ABS(" + formula[0] + ")-ABS(" + formula[1] + ")";
                                else
                                    Test_Range.Formula = "=(" + formula[0] + " - " + formula[1] + ")";

                                _Cruce.Diferencia = Test_Range.get_Value(Type.Missing).ToString();
                                xlSht.Cells[3, 1] = ValorAnterior;// restauro
                            }

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
                                if ((_Cruce.Diferencia == "") || (_Cruce.Diferencia == null))
                                {
                                    _Cruce.Diferencia = "0";
                                }

                                if (_Cruce.Diferencia != "0") // puede ser negativa
                                {
                                    //calculo la diferencia
                                    if (_Cruce.FormulaExcel.Contains("="))
                                    {
                                        if (_Cruce.FormulaExcel.Contains("="))
                                        {
                                            formula = _Cruce.FormulaExcel.Split('=');
                                            Test_Range = (Range)xlSht.get_Range("A4");
                                            ValorAnterior = Test_Range.get_Value(Type.Missing);
                                            Test_Range.Formula = "=" + formula[0];
                                            _Cruce.Grupo1 = Test_Range.get_Value(Type.Missing).ToString();
                                            xlSht.Cells[4, 1] = ValorAnterior;// restauro

                                            Test_Range = (Range)xlSht.get_Range("A5");
                                            ValorAnterior = Test_Range.get_Value(Type.Missing);
                                            Test_Range.Formula = "=" + formula[1];
                                            _Cruce.Grupo2 = Test_Range.get_Value(Type.Missing).ToString();
                                            xlSht.Cells[5, 1] = ValorAnterior;// restauro
                                        }
                                    }
                                    Globals.ThisAddIn._result.Add(_Cruce);
                                }
                            }

                            if (progress <= 70)
                            {
                                progress += 10;
                                pgbCruces.Value = progress;
                            }
                        }

                        Generales.Proteccion(true);
                    }
                    catch (Exception ex)
                    {
                        Generales.Proteccion(true);
                        MessageBox.Show(ex.Message, "Cruces", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }
                    progress += 5;

                    pgbCruces.Value = progress;
                }
                else if (_TemplateType == null)
                {
                    MessageBox.Show("Archivo no válido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            progress += 15;
            pgbCruces.Value = progress;
            if (Globals.ThisAddIn._result.Count > 0)
            {
                Globals.ThisAddIn.TaskPane.Visible = true;
                FIllValidacionDeCruceUC(Globals.ThisAddIn._result.ToArray());
                CreatePDF(Globals.ThisAddIn._result.ToArray(), _Cruces, _Path, _Excel.Name);
            }
            else
            {
                MessageBox.Show("No se encontraron diferencias", "Información Correcta", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            this.Hide();
        }
        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
        private void CreatePDF(oCruce[] _result, oCruce[] cruces, string path, string NombreLibro)
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
            //doc.Add(new Paragraph("eISSIF XML 17"));
            //doc.Add(new Paragraph("Cruces", _standardFont));
            //doc.Add(new Paragraph("SIPRED - ESTADOS FINANCIEROS GENERAL"));

            var titulo = new Paragraph("eISSIF XML 17", titlefont);
            titulo.Alignment = Element.ALIGN_CENTER;
            doc.Add(titulo);

            titulo = new Paragraph(NombreLibro, titlefont);
            titulo.Alignment = Element.ALIGN_CENTER;
            doc.Add(titulo);

            titulo = new Paragraph("Informe de Cruces: Diferencias", titlefont);
            titulo.Alignment = Element.ALIGN_CENTER;
            doc.Add(titulo);

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
                    PdfPCell cellcolumna = new PdfPCell(new Phrase(Generales.ColumnAdress(detail.Columna), _standardFont));
                    cellcolumna.BorderWidth = 0;
                    cellcolumna.BackgroundColor = new BaseColor(color);
                    PdfPCell cellconceptodet = new PdfPCell(new Phrase(detail.Concepto, _standardFont));
                    cellconceptodet.BorderWidth = 0;
                    cellconceptodet.BackgroundColor = new BaseColor(color);
                    cellconceptodet.Colspan = 2;

                    var strgpo1 = string.Empty;
                    var strgpo2 = string.Empty;

                    if (detail.Original != "")
                    {
                        if (formula1.Contains(detail.Original))
                            strgpo1 = detail.Valor == "0" ? "" : detail.Valor;

                        if (formula2.Contains(detail.Original))
                            strgpo2 = detail.Valor == "0" ? "" : detail.Valor;
                    }
                    else
                    {
                        if (detail.Grupo == 0)
                            strgpo1 = detail.Valor == "0" ? "" : detail.Valor;
                        else
                          if (detail.Grupo == 1)
                            strgpo2 = detail.Valor == "0" ? "" : detail.Valor;
                    }
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

                PdfPCell cellcalc = new PdfPCell(new Phrase("Cálculos", _standardFontbold));
                cellcalc.BorderWidth = 0;
                cellcalc.HorizontalAlignment = Element.ALIGN_RIGHT;
                cellcalc.Colspan = 5;

                PdfPCell cellgpot1 = new PdfPCell(new Phrase(item.Grupo1, _standardFont));
                cellgpot1.BorderWidth = 0;
                cellgpot1.HorizontalAlignment = Element.ALIGN_RIGHT;

                PdfPCell cellgpot2 = new PdfPCell(new Phrase(item.Grupo2, _standardFont));
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

                PdfPCell celldif = new PdfPCell(new Phrase(item.Diferencia, _standardFontbold));
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

            Process.Start(filepath);
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
        /// <summary>
        /// 
        /// Método que llena las tablas, textbox y lista del panel de valiadacion de los cruces
        /// El parámetro "_result" ya viene con todas las validaciones hechas.
        /// 
        /// </summary>
        /// <param name="_result"></param>
        private void FIllValidacionDeCruceUC(oCruce[] _result)
        {
            Globals.ThisAddIn.vdcUserControl.lst_Anexos.Items.Clear();
            Globals.ThisAddIn.vdcUserControl.dgv_DiferenciasEnCruces.DataSource = null;
            Globals.ThisAddIn.vdcUserControl.dgv_DiferenciasEnCruces.Rows.Clear();
            Globals.ThisAddIn.vdcUserControl.dgv_LadoDerechoDeFormula.DataSource = null;
            Globals.ThisAddIn.vdcUserControl.dgv_LadoDerechoDeFormula.Rows.Clear();
            Globals.ThisAddIn.vdcUserControl.dgv_LadoIzquierdoDeFormula.DataSource = null;
            Globals.ThisAddIn.vdcUserControl.dgv_LadoIzquierdoDeFormula.DataSource = null;
            Globals.ThisAddIn.vdcUserControl.txt_CrucesConDiferencia.Text = "0";
            Globals.ThisAddIn.vdcUserControl.txt_SumTotalLadoDerecho.Text = "0";
            Globals.ThisAddIn.vdcUserControl.txt_SumTotalLadoDerecho.Text = "0";
            Globals.ThisAddIn.vdcUserControl.txt_SumTotalLadoIzquierdo.Text = "0";
            Globals.ThisAddIn.vdcUserControl.txt_Formula.Text = "";
            
            var _TodosLosAnexos = (from items in _result
                                   from details in items.CeldasFormula
                                   orderby Int16.Parse(details.Anexo.Substring(6))
                                   select details.Anexo)
                                   .Distinct().ToList();

            foreach (var a in _TodosLosAnexos)
            {
                Globals.ThisAddIn.vdcUserControl.lst_Anexos.Items.Add(a);
            }

            if (Globals.ThisAddIn.vdcUserControl.lst_Anexos.Items.Count > 0)
            {
                Globals.ThisAddIn.vdcUserControl.lst_Anexos.SelectedIndex = 0;
            }

            Globals.ThisAddIn.vdcUserControl.txt_CrucesConDiferencia.Text = Globals.ThisAddIn._result.Count().ToString();
            Globals.ThisAddIn.vdcUserControl.txt_TotalCruces.Text = _TotalValidaciones.ToString();
        }
    }
}