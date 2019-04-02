using ExcelAddIn.Objects;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelAddIn1
{
    public partial class frmInfomeDeVerificaciones : Form
    {
        oCruce[] _CrucesLocalGbl=null;
        public frmInfomeDeVerificaciones()
        {
            InitializeComponent();
        }

        private void frmInfomeDeVerificaciones_Load(object sender, EventArgs e)
        {
            txt_TotalCruces.Text = "0";
            txt_TotalCrucesProcesados.Text = "0";
            txt_TotalLadoDer.Text = "0";
            txt_TotalLadoIzq.Text = "0";
            cmb_Vista.SelectedIndex = 0;
        }

        private void mItem_Detalle_Click(object sender, EventArgs e)
        {
            try
            {
                FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);

                if (_CrucesLocalGbl != null && _CrucesLocalGbl.Count() > 0)
                {
                    CreatePDF(_CrucesLocalGbl.ToArray(), Globals.ThisAddIn._TotalCruces, ExcelAddIn.Access.Configuration.Path, _Excel.Name);
                }
                else
                {
                    MessageBox.Show("No hay datos para crear el archivo PDF", "Sin Datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al crear el archivo PDF: {ex.Message.ToString()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            
        }

        private void cmb_Vista_SelectedIndexChanged(object sender, EventArgs e)
        {
            int _cmbIndice = cmb_Vista.SelectedIndex;

            txt_TotalCruces.Text = "";
            txt_TotalCrucesProcesados.Text = "";
            dgv_Cruce.DataSource = null;
            dgv_Cruce.Rows.Clear();
            txt_Formulas.Text = "";
            dgv_Indice.DataSource = null;
            dgv_Indice.Rows.Clear();
            txt_TotalLadoIzq.Text = "";
            txt_TotalLadoDer.Text = "";

            try
            {
                VerificaVista(_cmbIndice);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message.ToString()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgv_Cruce_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int _IdCruce = Convert.ToInt16(dgv_Cruce.CurrentRow.Cells[0].Value.ToString());
            FillDGIndiceConcepto(_CrucesLocalGbl, _IdCruce);

        }

        private void dgv_Indice_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //int _IdCruce = Convert.ToInt16(dgv_Cruce.CurrentRow.Cells[0].Value.ToString());
            //FillDGIndiceConcepto(_CrucesLocalGbl, _IdCruce);

        }

        #region Métodos y Funciones
        private void VerificaVista(int _cmbIndice)
        {
            try
            {
                switch (_cmbIndice)
                {
                    case 0:
                        _CrucesLocalGbl = Globals.ThisAddIn._result.ToArray();
                        FillDGConceptoCruce(_CrucesLocalGbl);
                        break;
                    case 1:
                        _CrucesLocalGbl = Globals.ThisAddIn._CrucesSinDiferencia.ToArray();
                        FillDGConceptoCruce(_CrucesLocalGbl);
                        break;
                    case 2:
                        _CrucesLocalGbl = Globals.ThisAddIn._CrucesQueNoAplican.ToArray();
                        FillDGConceptoCruce(_CrucesLocalGbl);
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"[frmInformeDeverifiaciones].[VerificaVista].[60]=> {ex.Message.ToString()}");
            }
        }

        private void FillDGConceptoCruce(oCruce[] _Cruces)
        {
            try
            {
                int count = _Cruces.Count();
                int _IdCruce;

                if (_Cruces.Count() > 0)
                {
                    bool _TieneNota;
                    foreach (var item in _Cruces)
                    {
                        _TieneNota = false;
                        if (!item.Nota.ToString().Equals(""))
                        {
                            _TieneNota = true;
                        }
                        dgv_Cruce.Rows.Add(item.IdCruce.ToString(), item.Concepto, item.Diferencia, item.Nota, _TieneNota, "");
                    }

                    txt_TotalCruces.Text = Globals.ThisAddIn._TotalCruces.Count().ToString();
                    txt_TotalCrucesProcesados.Text = _Cruces.Count().ToString();
                    _IdCruce = Convert.ToInt16(dgv_Cruce.CurrentRow.Cells[0].Value.ToString());
                    FillDGIndiceConcepto(_Cruces, _IdCruce);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"[frmInformeDeverifiaciones].[FillDGConceptoCruce].[94]=> {ex.Message.ToString()}");
            }
        }

        private void FillDGIndiceConcepto(oCruce[] _Cruces, int _IdCruce)
        {
            dgv_Indice.DataSource = null;
            dgv_Indice.Rows.Clear();
            txt_Formulas.Text = "";
            txt_TotalLadoDer.Text = "";
            txt_TotalLadoIzq.Text = "";

            try
            {
                int _SumLI = 0;
                int _SumLD =0;
                string _Gpo1;
                string _Gpo2;

                if (_Cruces.Count() > 0)
                {
                    var _Indices = (from item in _Cruces
                                   from details in item.CeldasFormula
                                   where item.IdCruce == _IdCruce
                                   select new
                                   {
                                       item.Formula,
                                       item.Condicion,
                                       details.Original,
                                       details.Indice,
                                       details.Concepto,
                                       details.Columna,
                                       details.Valor
                                   }).ToList();

                    foreach (var index in _Indices)
                    {
                        _Gpo1 = "0";
                        _Gpo2 = "0";
                        String[] _SplitFormula = index.Formula.Split('=');

                        if (_SplitFormula[0].Contains(index.Original.ToString()))
                        {
                            if (index.Valor.Equals("") || String.IsNullOrEmpty(index.Valor) || String.IsNullOrWhiteSpace(index.Valor))
                            {
                                _Gpo1 = "0";
                            }
                            else
                            {
                                _Gpo1 = index.Valor;
                            }                            
                            _SumLI = _SumLI + Convert.ToInt32(_Gpo1);                           
                        }

                        if (_SplitFormula[1].Contains(index.Original.ToString()))
                        {
                            if (index.Valor.Equals("") || String.IsNullOrEmpty(index.Valor) || String.IsNullOrWhiteSpace(index.Valor))
                            {
                                _Gpo2 = "0";
                            }
                            else
                            {
                                _Gpo2 = index.Valor;
                            }                            
                            _SumLD = _SumLD + Convert.ToInt32(_Gpo2);                            
                        }    
                                                
                        dgv_Indice.Rows.Add(index.Indice, index.Concepto, index.Columna, _Gpo1, _Gpo2);                      
                    }

                    txt_Formulas.Text = _Indices[0].Formula;
                    txt_Formulas.AppendText("\r\n" + _Indices[0].Condicion);
                    txt_TotalLadoIzq.Text = _SumLI.ToString();
                    txt_TotalLadoDer.Text = _SumLD.ToString();
                }                               
            }
            catch (Exception ex)
            {
                throw new Exception($"[frmInformeDeverifiaciones].[FillDGIndiceConcepto].[123]=> {ex.Message.ToString()}");
            }

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


        #endregion

        private void mItem_Cerrar_Click(object sender, EventArgs e)
        {
            this.Close();
            this.Dispose();
        }
    }
}
