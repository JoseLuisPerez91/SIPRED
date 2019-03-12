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
using ExcelAddIn.Logic;
using ExcelAddIn.Access;

namespace ExcelAddIn1 {
    public partial class Ribbon2
    {
        #region variable
        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
        string[,] HojasSPR = new string[,] {
                {"Contribuyente".ToUpper()          , "31"  ,"3"    ,""                     },
                {"Contador".ToUpper()               , "35"  ,"3"    ,""                     },
                {"Representante".ToUpper()          , "36"  ,"3"    ,""                     },
                {"Generales".ToUpper()              , "446" ,"3"    ,""                     },
                {"Anexo 1".ToUpper()                , "0"   ,"10"   ,""                     },
                {"Anexo 2".ToUpper()                , "0"   ,"9"    ,""                     },
                {"Anexo 3".ToUpper()                , "0"   ,"22"   ,""                     },
                {"Anexo 4".ToUpper()                , "0"   ,"5"    ,""                     },
                {"Anexo 5".ToUpper()                , "0"   ,"14"   ,""                     },
                {"Anexo 6".ToUpper()                , "0"   ,"5"    ,"Generales|C34"        },
                {"Anexo 7".ToUpper()                , "0"   ,"37"   ,""                     },
                {"Anexo 8".ToUpper()                , "0"   ,"9"    ,""                     },
                {"Anexo 9".ToUpper()                , "0"   ,"9"    ,""                     },
                {"Anexo 10".ToUpper()               , "0"   ,"15"   ,""                     },
                {"Anexo 11".ToUpper()               , "0"   ,"4"    ,""                     },
                {"Anexo 12".ToUpper()               , "0"   ,"13"   ,"Generales|C96"        },
                {"Anexo 13".ToUpper()               , "0"   ,"10"   ,"Generales|C97"        },
                {"Anexo 14".ToUpper()               , "0"   ,"12"   ,""                     },
                {"Anexo 15".ToUpper()               , "0"   ,"4"    ,""                     },
                {"Anexo 16".ToUpper()               , "0"   ,"11"   ,"Generales|C57"        },
                {"Anexo 17".ToUpper()               , "0"   ,"4"    ,"Generales|C57"        },
                {"Anexo 18".ToUpper()               , "0"   ,"4"    ,""                     },
                {"Anexo 19".ToUpper()               , "0"   ,"7"    ,"Generales|C98"        },
                {"Anexo 20".ToUpper()               , "0"   ,"9"    ,""                     },
                {"Anexo 21".ToUpper()               , "0"   ,"12"   ,"Generales|C100"       },
                {"Anexo 22".ToUpper()               , "0"   ,"25"   ,"Generales|C101"       },
                {"Anexo 23".ToUpper()               , "0"   ,"14"   ,"Generales|C61,C62"    },
                {"CDF".ToUpper()                    , "78"  ,"5"    ,""                     },
                {"MPT".ToUpper()                    , "111" ,"3"    ,""                     },
                {"Notas".ToUpper()                  , "48"  ,"1"    ,""                     },
                {"Declaratoria".ToUpper()           , "45"  ,"1"    ,""                     },
                {"Opinión".ToUpper()                , "45"  ,"1"    ,""                     },
                {"Informe".ToUpper()                , "45"  ,"1"    ,""                     },
                {"Información Adicional".ToUpper()  , "45"  ,"1"    ,""                     }
            };
        String[] nombre;
        #endregion
        #region metodos
        public void MensageBloqueo(Excel.Worksheet Sh)
        {
            String CondCad = "";
            string[] arg;
            string[] cond;
            Boolean res = true;
            String Vcon = "";
            //cargar array de nombres
            Cargararraynombre(HojasSPR);

            String nom = Sh.Name.ToString().Trim();
            int ind = Array.IndexOf(nombre, Sh.Name.ToString().Trim().ToUpper());
            if (ind != -1)
            {
                //ind++;
                Sh.Activate();
                if (HojasSPR[ind, 3].Trim().Length > 0)
                {
                    //Capturo la condicion
                    CondCad = HojasSPR[ind, 3].Trim();
                    arg = CondCad.Split('|');
                    nom = arg[0].ToString().Trim();
                    ind = Array.IndexOf(nombre, nom.ToUpper());
                    cond = arg[1].ToString().Trim().Split(',');

                    foreach (string i in cond)
                    {
                        Vcon = ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets["Generales"]).Range[i].Formula;

                        if (Vcon.Trim().ToUpper().Contains("SI"))
                        {
                            res = false;
                            break;
                        }
                        if (Vcon.Trim().ToUpper().Contains("NO"))
                        {
                            res = true;
                        }
                    }
                    if (res)
                    {
                        MessageBox.Show("No es posible seleccionar el anexo debido a que se encuentra deshabilitado.", "SPRIND", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        ((Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets[Sh.Index - 1]).Activate();
                    }
                }
            }
        }
        public void Cargararraynombre(string[,] val)
            {
                int numf = (val.Length) / val.GetLength(1);
                nombre = new String[numf];
                for (int k = 0; k < numf; k++)
                {
                    nombre[k] = val[k, 0];
                }
            }
            public void GuardarExcel()
            {
                //guardar nuevo libro
                object obj = Type.Missing;
                Excel.Workbook libron = Globals.ThisAddIn.Application.ActiveWorkbook;
                SaveFileDialog1 = new SaveFileDialog()
                {
                    DefaultExt = "*.xlsx",
                    //SaveFileDialog1.FileName = Globals.ThisAddIn.Application.ActiveWorkbook.Name + ".xls";
                    FileName = libron.Name + ".xlsx",
                    Filter = "Archivos de Excel (*.xlsx)|*.xlsx"
                };
                if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    libron.SaveAs(SaveFileDialog1.FileName, Excel.XlFileFormat.xlOpenXMLWorkbook, obj, obj, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlUserResolution, true, obj, obj, obj);
                }
            }
        #endregion

        bool _Connection = true; //new lSerializados().CheckConnection(Configuration.UrlConnection);
        string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
        string _Title = "Conexión de Red";
        int NroFilaPrincipal = 0;
        int NroColPrincipal = 0;
        bool tieneformula = false;
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e) {

        }
        public void btnNew_Click(object sender, RibbonControlEventArgs e)
        {
            if (_Connection)
            {
                Nuevo _New = new Nuevo();
                _New.Show();
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void btnCruces_Click(object sender, RibbonControlEventArgs e)
        {
            if (_Connection)
            {
                Cruce _Cruce = new Cruce();
                _Cruce.Show();
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void btnPlantilla_Click(object sender, RibbonControlEventArgs e)
        {
            if (_Connection)
            {
                LoadTemplate _Template = new LoadTemplate();
                _Template.Show();
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            tieneformula = false;
            string tag = "";

            try
            {
                objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal, 1];
                IndicePrevio = objRange.get_Value(Type.Missing).ToString();
                if (IndicePrevio.ToUpper().Trim() != "EXPLICACION")
                {
                    foreach (Excel.Name item in wb.Names)
                    {
                        if (item.Name.Substring(0, 3) == "IA_")
                        {
                            tag = item.RefersToRange.Cells.get_Address();

                            if (tag == objRange.Address)
                            {

                                if ((NroFilaPrincipal - 1) > 0)
                                {
                                    var RangeConFr = ActiveWorksheet.get_Range(string.Format("{0}:{0}", NroFilaPrincipal, Type.Missing));
                                    iTotalColumns = ActiveWorksheet.UsedRange.Columns.Count;

                                    while (k <= iTotalColumns)
                                    {
                                        if (RangeConFr.Cells[k].HasFormula)
                                        {
                                            tieneformula = true;
                                            break;
                                        }

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
                        string indicex = "";

                        if (ConceptoPrevio != null)
                        {
                            ConceptoPrevio = ConceptoPrevio.ToString();
                            CncValido = Generales.EsConceptoValido(ConceptoPrevio);

                            if (CncValido)
                            {
                                NroFilaPrincipal = objRange.Row;
                                NroColPrincipal = objRange.Column;
                                if ((NroFilaPrincipal - 1) > 0)
                                {
                                    var RangeConFr = ActiveWorksheet.get_Range(string.Format("{0}:{0}", NroFilaPrincipal - 1, Type.Missing));
                                    objRange = (Excel.Range)ActiveWorksheet.Cells[NroFilaPrincipal - 1, 1];
                                    if (objRange.get_Value(Type.Missing) != null)
                                        indicex = objRange.get_Value(Type.Missing).ToString();

                                    if (indicex != "01060025000000")
                                    {
                                        iTotalColumns = ActiveWorksheet.UsedRange.Columns.Count;

                                        while (k <= iTotalColumns)
                                        {
                                            if (RangeConFr.Cells[k].HasFormula)
                                            {
                                                tieneformula = true;
                                                break;
                                            }

                                            k = k + 1;
                                        }
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
                    sheet.Unprotect(ExcelAddIn.Access.Configuration.PwsExcel);
                    currentCell = (Excel.Range)Globals.ThisAddIn.Application.Selection;
                    objRange = (Excel.Range)sheet.Cells[currentCell.Cells.Row + 1, 1];
                    IndiceSiguiente = objRange.Value2;

                    if (IndiceSiguiente != null)
                        if (IndiceSiguiente.ToUpper().Trim() == "EXPLICACION")
                            objRange.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                    currentCell.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    NombreRangosDEL.Sort();
                    string NM = NombreRangosDEL.FirstOrDefault();
                    sheetControl.Controls.Remove(NM);

                    foreach (Excel.Name item2 in wb.Names)
                    {
                        if (item2.Name.Substring(0, 3) == "IA_")
                        {
                            NombreRangos.Add(item2.Name);
                        }
                    }

                    string[] split = NM.Split('_');
                    NM = split[1];
                    // foreach (string Nm in NombreRangosDEL)
                    long NamedRng = Convert.ToInt64(NM) + 100;
                    string IndiceSig = "0" + Convert.ToString(NamedRng);
                    string IndiceAnt = "";

                    while (NombreRangos.Contains("IA_" + IndiceSig))
                    {
                        sheetControl.Controls.Remove("IA_" + IndiceSig);

                        NamedRng = Convert.ToInt64(IndiceSig) + 100;
                        IndiceSig = "0" + Convert.ToString(NamedRng);
                    }

                    FilaPadre.Sort();
                    int row = FilaPadre.FirstOrDefault();

                    objRange = (Excel.Range)sheet.Cells[row, 1];

                    if (objRange.get_Value(Type.Missing) != null)
                        IndiceActivo = objRange.get_Value(Type.Missing).ToString();

                    objRange = (Excel.Range)sheet.Cells[row - 1, 1];

                    if (objRange.get_Value(Type.Missing) != null)
                        IndiceAnt = objRange.get_Value(Type.Missing).ToString();
                    
                    //me salto la explciacion
                    if (IndiceAnt.Trim() == "EXPLICACION")
                    {
                        objRange = (Excel.Range)sheet.Cells[row - 2, 1];
                        if (objRange.get_Value(Type.Missing) != null)
                            IndiceAnt = objRange.get_Value(Type.Missing).ToString();
                    }

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

                        if (objRange.get_Value(Type.Missing) != null)
                            IndiceActivo = objRange.get_Value(Type.Missing).ToString();
                        else
                            break;
                    }

                    row = Generales.DameRangoPrincipal(FilaPadre.FirstOrDefault(), sheet);// busco el numero de fila OTRO para agregarle luego la sumatoria de los indices nuevos
                    Excel.Range objRangeJ = ((Excel.Range)sheet.Cells[FilaPadre[0], 1]);
                    objRangeJ.Select();

                    try
                    { // limpio si hay error en la formula
                        Excel.Range objRangeI = ((Excel.Range)sheet.Cells[row, 1]).SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors);//obten las celdas con errores
                        string NombreHoja = sheet.Name.ToUpper().Replace(" ", "");
                        List<oSubtotal> ColumnasST = Generales.DameColumnasST(NombreHoja);

                        foreach (oSubtotal ST in ColumnasST)
                        {
                            objRangeI = sheet.get_Range(ST.Columna + row.ToString(), ST.Columna + row.ToString());
                            objRangeI.Clear();
                        }
                    }
                    catch (Exception ex)
                    {
                    }

                    sheet.Protect(ExcelAddIn.Access.Configuration.PwsExcel, true, true, false, true, true, true, true, false, false, false, false, false, false, true, false);
                }
            }
            catch (Exception ex)
            {
                //  MessageBox.Show(ex.Message);
            }
        }
        private void btnAgregarExplicacion_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Excel.Worksheet ActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

                Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell; // Fila activa
                Excel.Range objRange = (Excel.Range)ActiveWorksheet.Cells[currentCell.Row, 9];// voy a la columna 9
                string DebeExplicar = objRange.get_Value(Type.Missing);
                objRange = (Excel.Range)ActiveWorksheet.Cells[currentCell.Row, 1]; // me aseguro que sea la activa en la columna 1
                string Indice = objRange.get_Value(Type.Missing);
                objRange = (Excel.Range)ActiveWorksheet.Cells[currentCell.Row, 2];// voy a la columna 2 de concepto del indice activo
                string Concepto = objRange.get_Value(Type.Missing);
                objRange = (Excel.Range)ActiveWorksheet.Cells[currentCell.Row + 1, 1]; // voy al indice siguiente
                string IndiceSig = objRange.get_Value(Type.Missing);

                if (Indice != null)
                {
                    if (Indice.ToString().ToUpper().Trim() == "EXPLICACION")
                        MessageBox.Show("El índice " + Indice.ToString() + "no es válido", "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    else
                    if (DebeExplicar != null)
                    {
                        if (DebeExplicar.ToString().ToUpper() == "SI")
                        {

                            if (IndiceSig != null)
                            {
                                if (IndiceSig.ToString().ToUpper().Trim() == "EXPLICACION")
                                    MessageBox.Show("El índice " + Indice.ToString() + " ya tiene una explicación asociada", "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                else
                                    CargaFormulario(Indice, Concepto);


                            }
                            else
                                CargaFormulario(Indice, Concepto);


                        }
                        else
                            MessageBox.Show("Debe haber una respuesta afirmativa en la columna ' EXPLICAR VARIACION ' para agregar una explicación en el índice " + Indice.ToString(), "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                        MessageBox.Show("Debe haber una respuesta afirmativa en la columna ' EXPLICAR VARIACION ' para agregar una explicación en el índice " + Indice.ToString(), "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                    MessageBox.Show("El índice no es válido ", "Agregar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch { }
        }
        public static void CargaFormulario(string Indice, string Concepto)
        {
            Explicaciones NewExplicacion = new Explicaciones();
            NewExplicacion.Text = "Explicación índice " + Indice.ToString();
            if (Concepto != null)
                NewExplicacion.Text = NewExplicacion.Text + " " + Concepto.ToString();

            NewExplicacion.ShowDialog();
        }
        private void btnEliminaeExplicacion_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell;
            int NroRow = currentCell.Row;
            Excel.Worksheet NewActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            currentCell = (Excel.Range)NewActiveWorksheet.Cells[NroRow, 1];

            string indice = currentCell.Value2;
            if (indice.ToUpper().Trim() == "EXPLICACION")
            {
                NewActiveWorksheet.Unprotect(ExcelAddIn.Access.Configuration.PwsExcel);

                currentCell.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

                NewActiveWorksheet.Protect(ExcelAddIn.Access.Configuration.PwsExcel, true, true, false, true, true, true, true, false, false, false, false, false, false, true, false);
            }
            else
                MessageBox.Show("La fila seleccionada no es una explicación ", "Eliminar Explicación", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        private void btnConvertir_Click(object sender, RibbonControlEventArgs e)
        {
        }
        private void btnTransferir_Click(object sender, RibbonControlEventArgs e)
        {
            if (_Connection)
            {
                FormulasComprobaciones form = new FormulasComprobaciones();
                form._formulas = false;
                form.ShowDialog();
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCrucesAdmin_Click(object sender, RibbonControlEventArgs e)
        {
            if (_Connection)
            {
                CrucesAdmin form = new CrucesAdmin();
                form.ShowDialog();
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnComprobacionesAdmin_Click(object sender, RibbonControlEventArgs e)
        {
            if (_Connection)
            {

                ComprobacionesAdmin form = new ComprobacionesAdmin();
                form.ShowDialog();
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnImprimir_Click(object sender, RibbonControlEventArgs e)
        {
            if (_Connection)
            {
                frmPreImprimir form = new frmPreImprimir();
                form.ShowDialog();
                var addIn = Globals.ThisAddIn;
                addIn.Imprimir();
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}