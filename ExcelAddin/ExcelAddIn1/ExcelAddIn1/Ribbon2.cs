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
using Microsoft.Win32;
using System.Data.SqlClient;
using System.Data;

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
        public void btnOpen_Click(object sender, RibbonControlEventArgs e)
        {
            if (_Connection)
            {
                OpenFileDialog _Abrir = new OpenFileDialog();

                _Abrir.Filter = "Archivo xlsm (*.xlsm)|*.xlsm";
                _Abrir.Title = "Abrir archivo xlsm";
                _Abrir.ShowDialog();

                if(_Abrir.FileName == "")
                {
                    MessageBox.Show("Debe especificar un archivo xlsm", "Archivo xlsm Invalido", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                const string userRoot = "HKEY_CURRENT_USER";
                const string subkey = "Software\\Microsoft\\Office\\Excel\\Addins\\SAT.Dictamenes.SIPRED.Client";
                const string keyName = userRoot + "\\" + subkey;
                object addInName = "SAT.Dictamenes.SIPRED.Client";
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

                Registry.SetValue(keyName, "LoadBehavior", 0);
                Globals.ThisAddIn.Application.COMAddIns.Item(ref addInName).Connect = false;
                //try
                //{
                //    wb.Close();
                //}
                //catch { }
                Globals.ThisAddIn.Application.Visible = true;
                Globals.ThisAddIn.Application.Workbooks.Open(_Abrir.FileName);

                //Libro Actual de Excel.
                Excel.Worksheet xlSht;
                wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                int count = wb.Worksheets.Count;
                bool _existeS = false;
                bool _existeO = false;

                for(int _wCount = 1; _wCount<= count; _wCount++)
                {
                    string _sName = wb.Worksheets[_wCount].Name;

                    if(_sName == "SIPRED")
                    {
                        _existeS = true;
                    }
                    if (_sName == "_Open")
                    {
                        _existeO = true;
                    }
                }

                if (!_existeS)
                {
                    xlSht = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, wb.Worksheets[count], Type.Missing, Type.Missing);
                    xlSht.Name = "SIPRED";
                }
                if (!_existeO)
                {
                    count = wb.Worksheets.Count;
                    xlSht = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, wb.Worksheets[count], Type.Missing, Type.Missing);
                    xlSht.Name = "_Open";
                }
                wb.Save();

                FormulasComprobaciones _Formulas = new FormulasComprobaciones();
                _Formulas._Form = null;
                _Formulas._Template = null;
                _Formulas._Tipo = "";
                _Formulas._formulas = true;
                _Formulas._Open = true;
                _Formulas.ShowDialog();
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
            //string _Path = ExcelAddIn.Access.Configuration.Path;

            //var fecha = DateTime.Now;
            //var name = "Cruce_" + fecha.Year.ToString() + fecha.Month.ToString() + fecha.Day.ToString() + fecha.Hour.ToString() + fecha.Minute.ToString() + fecha.Second.ToString();
            //var filepath = _Path + "\\" + name + ".pdf";
            //// Creamos el documento con el tamaño de página tradicional
            //Document doc = new Document(PageSize.LETTER);
            //// Indicamos donde vamos a guardar el documento
            //PdfWriter writer = PdfWriter.GetInstance(doc, new FileStream(filepath, FileMode.Create));
            //// Le colocamos el título y el autor
            //doc.AddTitle("Cruces");
            //doc.AddCreator("S-DAT");
            //// Abrimos el archivo
            //doc.Open();
            //PdfPTable tabla = new PdfPTable(3);

            //for (int i = 0; i < 15; i++)
            //{
            //    tabla.AddCell("A " + i);
            //    tabla.AddCell("B " + i);
            //}
            //doc.Add(tabla);

            //doc.Close();
            //writer.Close();
            string _CnStr = string.Format(Configuration.ConnectionStringPrellenado, Configuration.Server, Configuration.DataBase, Configuration.User, Configuration.Password);
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);

            string _RFC = "CVG080811RU4";
            int _Anio = 2018;

            DataTable dt = new DataTable();
            dt.Clear();
            dt.Columns.Add("INDICE");
            dt.Columns.Add("SALDO");
            dt.Columns.Add("CUENTAS");

            DataTable dt2 = new DataTable();
            dt2.Clear();
            dt2.Columns.Add("INDICE");
            dt2.Columns.Add("SALDO");
            dt2.Columns.Add("CUENTAS");

            SqlConnection _DbConn = new SqlConnection(_CnStr);

            try
            {
                if (_Connection)
                {
                    using (_DbConn)
                    {
                        SqlCommand _SqlComm = new SqlCommand("dbo.SP_DAgrupa_ObtieneSaldoIndice", _DbConn);
                        _SqlComm.Parameters.AddWithValue("@RFC", _RFC);
                        _SqlComm.Parameters.AddWithValue("@Ejercicio", _Anio.ToString());
                        _SqlComm.Parameters.AddWithValue("@Indice", "");

                        _SqlComm.CommandType = CommandType.StoredProcedure;
                        SqlDataAdapter da = new SqlDataAdapter();

                        da.SelectCommand = _SqlComm;
                        da.SelectCommand.CommandType = CommandType.StoredProcedure;
                        da.Fill(dt);

                        _SqlComm.Parameters.Clear();

                        _SqlComm.Parameters.AddWithValue("@RFC", _RFC);
                        _SqlComm.Parameters.AddWithValue("@Ejercicio", (_Anio - 1).ToString());
                        _SqlComm.Parameters.AddWithValue("@Indice", "");
                        SqlDataAdapter daT2 = new SqlDataAdapter();
                        daT2.SelectCommand = _SqlComm;
                        daT2.SelectCommand.CommandType = CommandType.StoredProcedure;
                        daT2.Fill(dt2);
                    }

                }
                else
                {
                    MessageBox.Show("Porfavor verifique que tiene conexión a internet.", "Sin acceso a la red", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch (SqlException sqlex)
            {
                MessageBox.Show($"Error en la conexión. {sqlex.Message.ToString()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error en la conexión. {ex.Message.ToString()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                _DbConn.Close();

                if (dt.Rows.Count > 0 || dt2.Rows.Count > 0)
                {
                    Generales.Proteccion(false);
                    ValidaSaldoIndice(dt, dt2);
                    Generales.Proteccion(true);
                }
                else
                {
                    MessageBox.Show($"No hay datos para el cliente {_RFC}, periodo {_Anio}.", "Sin Datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

            }
        }
        private void btnEliminarIndice_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sheetControl = Globals.Factory.GetVstoObject(Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            string IndiceActivo = "";
            string IndiceSiguiente = "";
            string _NameFile = wb.Name;
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
                    int CantRowDelete = currentCell.Cells.Rows.Count;
                    objRange = (Excel.Range)sheet.Cells[currentCell.Cells.Row + 1, 1];
                    IndiceSiguiente = objRange.Value2;

                    if (IndiceSiguiente != null)
                    {
                        if (IndiceSiguiente.ToUpper().Trim() == "EXPLICACION")
                        {
                            objRange.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                            CantRowDelete += 1;
                        }
                    }

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
                        Excel.Range objRangeI = ((Excel.Range)sheet.Cells[row, 1]);//.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, Excel.XlSpecialCellsValue.xlErrors);//obten las celdas con errores
                        string NombreHoja = sheet.Name.ToUpper().Replace(" ", "");
                        List<oSubtotal> ColumnasST = Generales.DameColumnasST(NombreHoja);
                        int _Registro = 1;

                        foreach (oSubtotal ST in ColumnasST)
                        {
                            objRangeI = sheet.get_Range(ST.Columna + row.ToString(), ST.Columna + row.ToString());
                            //objRangeI.Clear();

                            if (_Registro == 1)
                            {
                                _Registro += 1;
                                Generales.ActualizarReferencia(_NameFile, sheet.Name.ToUpper(), ST.Columna + row.ToString(), NombreRangos.Count, ST.Columna, row.ToString(), CantRowDelete,"E");
                            }
                        }
                        //wb.Save();
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

                //ref
                string NombreHoja = NewActiveWorksheet.Name.ToUpper().Replace(" ", "");
                List<oSubtotal> ColumnasST = Generales.DameColumnasST(NombreHoja);
                int _Registro = 1;
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                string _NameFile = wb.Name;
             
                int row = Generales.DameRangoPrincipal(NroRow, NewActiveWorksheet);
                foreach (oSubtotal ST in ColumnasST)
                {
                    if (_Registro == 1)
                    {
                        _Registro += 1;
                        Generales.ActualizarReferencia(_NameFile, NewActiveWorksheet.Name.ToUpper(), ST.Columna + row.ToString(), 0, ST.Columna, row.ToString(), 1, "E");
                    }
                }
                //ref
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
                form._Open = false;
                form.TopMost = false;
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
        /// <summary>
        /// Validamos que el indice no venga vacio y tenga el formato correcto
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="dt2"></param>
        private void ValidaSaldoIndice(DataTable dt, DataTable dt2)
        {
            string _Indice;
            string _Saldo;
            string _Anexo;
            string _IndexNumber;

            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    _Anexo = "";
                    _Indice = row["INDICE"].ToString();
                    if (_Indice.ToString().Length == 14)
                    {

                        _Saldo = row["SALDO"].ToString();
                        _IndexNumber = _Indice.ToString().Substring(2, 2);
                        _Anexo = ObtenerStringAnexo(_IndexNumber);

                        LlenaCeldas(_Saldo, _Indice, _Anexo, "C");
                    }
                    else
                    {
                        continue;
                    }
                }
            }

            if (dt2.Rows.Count > 0)
            {
                foreach (DataRow row in dt2.Rows)
                {
                    _Anexo = "";
                    _Indice = row["INDICE"].ToString();
                    if (_Indice.ToString().Length == 14)
                    {

                        _Saldo = row["SALDO"].ToString();
                        _IndexNumber = _Indice.ToString().Substring(2, 2);
                        _Anexo = ObtenerStringAnexo(_IndexNumber);

                        LlenaCeldas(_Saldo, _Indice, _Anexo, "D");
                    }
                    else
                    {
                        continue;
                    }
                }
            }
        }

        private void LlenaCeldas(string _Saldo, string _Indice, string _Anexo, string _Columna)
        {
            int _MaxRow = 0;

            try
            {
                Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                Excel.Worksheet sheet = (Excel.Worksheet)wb.Worksheets.get_Item(_Anexo);
                _MaxRow = sheet.UsedRange.Count + 1;

                if (sheet != null)
                {
                    Excel.Range range = (Excel.Range)sheet.get_Range("A1:A" + _MaxRow.ToString());
                    Excel.Range findValue = range.Find(_Indice, Type.Missing, Excel.XlFindLookIn.xlValues,
                                                        Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows,
                                                        Excel.XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);
                    if (findValue != null)
                    {
                        range = (Excel.Range)sheet.Cells[findValue.Row, _Columna];
                        range.Value = _Saldo;
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al llenar la hoja {_Anexo}: {ex.Message.ToString()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string ObtenerStringAnexo(string _indexNumber)
        {
            string result = "";

            switch (_indexNumber)
            {
                case "01":
                    result = "Anexo 1";
                    break;
                case "02":
                    result = "Anexo 2";
                    break;
                case "03":
                    result = "Anexo 3";
                    break;
                case "04":
                    result = "Anexo 4";
                    break;
                case "05":
                    result = "Anexo 5";
                    break;
                case "06":
                    result = "Anexo 6";
                    break;
                case "07":
                    result = "Anexo 7";
                    break;
                case "08":
                    result = "Anexo 8";
                    break;
                case "09":
                    result = "Anexo 9";
                    break;
                case "10":
                    result = "Anexo 10";
                    break;
                case "11":
                    result = "Anexo 11";
                    break;
                case "12":
                    result = "Anexo 12";
                    break;
                case "13":
                    result = "Anexo 13";
                    break;
                case "14":
                    result = "Anexo 14";
                    break;
                case "15":
                    result = "Anexo 15";
                    break;
                case "16":
                    result = "Anexo 16";
                    break;
                case "17":
                    result = "Anexo 17";
                    break;
                case "18":
                    result = "Anexo 18";
                    break;
                case "19":
                    result = "Anexo 19";
                    break;
                case "20":
                    result = "Anexo 20";
                    break;
            }

            return result;
        }
    }
}