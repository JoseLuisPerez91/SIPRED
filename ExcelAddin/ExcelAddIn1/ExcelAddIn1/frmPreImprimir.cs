using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using ExcelAddIn.Access;
using System.IO;

namespace ExcelAddIn1
{
    public partial class frmPreImprimir : Form
    {
        public bool _ProcessJson;
        public frmPreImprimir()
        {
            InitializeComponent();
        }
        private void frmCarga_Load(object sender, EventArgs e)
        {
            string _Path = Configuration.Path;
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            this._ProcessJson = false;
            this.Visible = true;
            
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
                        this.btnAccept.Enabled = false;
                        return;
                    }
                    else
                    {
                        this._ProcessJson = true;
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

                this._ProcessJson = true;
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

            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
            
            if(_Excel.Extension != ".xlsm")
            {
                MessageBox.Show("Archivo no válido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;
            }

            QuitarFormulas();
        }
        private void frmCarga_Shown(object sender, EventArgs e)
        {
        }
        private void btnAccept_Click(object sender, EventArgs e)
        {
        }
        private void button1_Click(object sender, EventArgs e)
        {
        }
        #region Variables
        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

        string[,] HojasSPR = new string[,] {
            {"Contribuyente".ToUpper()          , "31"  ,"3"    ,""                     ,"Contribuyente"},
            {"Contador".ToUpper()               , "35"  ,"3"    ,""                     ,"Contador"},
            {"Representante".ToUpper()          , "36"  ,"3"    ,""                     ,"Representante"},
            {"Generales".ToUpper()              , "446" ,"3"    ,""                     ,"Generales"},
            {"Anexo 1".ToUpper()                , "0"   ,"10"   ,""                     ,"1.- ESTADO DE SITUACION FINANCIERA"},
            {"Anexo 2".ToUpper()                , "0"   ,"9"    ,""                     ,"2.- ESTADO DE RESULTADO INTEGRAL"},
            {"Anexo 3".ToUpper()                , "0"   ,"22"   ,""                     ,"3.- ESTADO DE CAMBIOS EN EL CAPITAL CONTABLE"},
            {"Anexo 4".ToUpper()                , "0"   ,"5"    ,""                     ,"4.- ESTADO DE FLUJOS DE EFECTIVO "},
            {"Anexo 5".ToUpper()                , "0"   ,"14"   ,""                     ,"5.- INTEGRACION ANALITICA DE VENTAS O INGRESOS NETOS "},
            {"Anexo 6".ToUpper()                , "0"   ,"5"    ,""                     ,"6.- DETERMINACION DEL COSTO DE LO VENDIDO PARA EFECTOS CONTABLES Y DEL IMPUESTO SOBRE LA RENTA "},
            {"Anexo 7".ToUpper()                , "0"   ,"37"   ,""                     ,"7.- ANALISIS COMPARATIVO DE LAS SUBCUENTAS DE GASTOS"},
            {"Anexo 8".ToUpper()                , "0"   ,"9"    ,""                     ,"8.- ANALISIS COMPARATIVO DE LAS SUBCUENTAS DEL RESULTADO INTEGRAL DE FINANCIAMIENTO"},
            {"Anexo 9".ToUpper()                , "0"   ,"9"    ,""                     ,"9.- RELACION DE CONTRIBUCIONES A CARGO DEL CONTRIBUYENTE COMO SUJETO DIRECTO O EN SU CARACTER DE RETENEDOR"},
            {"Anexo 10".ToUpper()               , "0"   ,"15"   ,""                     ,"10.- RELACION DE CONTRIBUCIONES POR PAGAR"},
            {"Anexo 11".ToUpper()               , "0"   ,"4"    ,""                     ,"11.- CONCILIACION ENTRE EL RESULTADO CONTABLE Y FISCAL PARA EFECTOS DEL IMPUESTO SOBRE LA RENTA"},
            {"Anexo 12".ToUpper()               , "0"   ,"13"   ,"Generales|C96"        ,"12.- OPERACIONES FINANCIERAS DERIVADAS CONTRATADAS CON RESIDENTES EN EL EXTRANJERO "},
            {"Anexo 13".ToUpper()               , "0"   ,"10"   ,"Generales|C97"        ,"13.- INVERSIONES PERMANENTES EN SUBSIDIARIAS, ASOCIADAS Y AFILIADAS RESIDENTES EN EL EXTRANJERO"},
            {"Anexo 14".ToUpper()               , "0"   ,"12"   ,""                     ,"14.- SOCIOS O ACCIONISTAS QUE TUVIERON ACCIONES O PARTES SOCIALES"},
            {"Anexo 15".ToUpper()               , "0"   ,"4"    ,""                     ,"15.- CONCILIACION ENTRE LOS INGRESOS DICTAMINADOS SEGUN ESTADO DE RESULTADO INTEGRAL Y LOS ACUMULABLES PARA EFECTOS DEL IMPUESTO SOBRE LA RENTA Y  EL TOTAL DE ACTOS O ACTIVIDADES PARA EFECTOS DEL IMPUESTO AL VALOR AGREGADO"},
            {"Anexo 16".ToUpper()               , "0"   ,"11"   ,"Generales|C57"        ,"16.- OPERACIONES CON PARTES RELACIONADAS"},
            {"Anexo 17".ToUpper()               , "0"   ,"4"    ,"Generales|C57"        ,"17.- INFORMACION DEL CONTRIBUYENTE SOBRE SUS OPERACIONES CON PARTES RELACIONADAS"},
            {"Anexo 18".ToUpper()               , "0"   ,"4"    ,""                     ,"18.- DATOS INFORMATIVOS "},
            {"Anexo 19".ToUpper()               , "0"   ,"7"    ,"Generales|C98"        ,"19.- INFORMACION DE LOS PAGOS REALIZADOS POR LA  DETERMINACION DEL IMPUESTO SOBRE LA RENTA E IMPUESTO AL ACTIVO DIFERIDO POR DESCONSOLIDACION AL 31 DE DICIEMBRE DE 2013 Y EL PAGADO HASTA EL 30 DE ABRIL DEL 2018"},
            {"Anexo 20".ToUpper()               , "0"   ,"9"    ,""                     ,"20.- INVERSIONES"},
            {"Anexo 21".ToUpper()               , "0"   ,"12"   ,"Generales|C100"       ,"21.- CUENTAS Y DOCUMENTOS POR COBRAR Y POR PAGAR EN MONEDA EXTRANJERA"},
            {"Anexo 22".ToUpper()               , "0"   ,"25"   ,"Generales|C101"       ,"22.- PRESTAMOS DEL EXTRANJERO "},
            {"Anexo 23".ToUpper()               , "0"   ,"14"   ,"Generales|C61,C62"    ,"23.- INTEGRACION DE PERDIDAS FISCALES DE EJERCICIOS ANTERIORES"},
            {"CDF".ToUpper()                    , "78"  ,"5"    ,""                     ,"CUESTIONARIO DE DIAGNOSTICO FISCAL (REVISION DEL CONTADOR PUBLICO)"},
            {"MPT".ToUpper()                    , "111" ,"3"    ,""                     ,"CUESTIONARIO EN MATERIA DE PRECIOS DE TRANSFERENCIA (REVISION DEL CONTADOR PÚBLICO)"},
            {"Notas".ToUpper()                  , "48"  ,"1"    ,""                     ,""},
            {"Declaratoria".ToUpper()           , "45"  ,"1"    ,""                     ,""},
            {"Opinión".ToUpper()                , "45"  ,"1"    ,""                     ,""},
            {"Informe".ToUpper()                , "45"  ,"1"    ,""                     ,""},
            {"Información Adicional".ToUpper()  , "45"  ,"1"    ,""                     ,""}
        };

        String[] nombre;
        #endregion
        public void QuitarFormulas()
        {
            //objeto vacio
            object obj = Type.Missing;
            int numhojas = 0;
            int EspacioFilas = 0;
            int EspacioColumnas = 0;
            int fila = 1;
            int columna = 1;
            int ind = 0;
            String nom = "";
            Range cell1;
            Range cell2;
            Range range;
            String psw = "";


            //Nuevo Excel
            Excel.Application exceln = new Excel.Application();
            //libro abierto
            Excel.Workbook libro = Globals.ThisAddIn.Application.ActiveWorkbook;
            //nuevo libro
            //Excel.Workbook libron = exceln.Workbooks.Add(obj);
            Excel.Workbook libron = libro;
            //libron = exceln.Workbooks.

            //obtenemos el numero de hojas
            numhojas = libro.Sheets.Count;
            //cargar array de nombres
            Cargararraynombre(HojasSPR);
            //seleccionamos instanci hoja
            Excel.Worksheet hoja = libro.Sheets[1];
            //Creamos instancia nueva hoja
            Excel.Worksheet hojan = new Excel.Worksheet();
            int[] hojas = new int[numhojas];
            //Contraseña
            //psw = "AAAABABABAAG";
            //LENARHOJAS
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            for (int i = 1; i < numhojas; i++)
            {
                //seleccionamos la hoja de orden i
                hoja = libro.Sheets[i];
                //Creamos nueva hoja de orden i
                hojan = libron.Sheets[i];
                hojas[i] = ValidarInt(Regex.Replace(hoja.Name, @"[^\d]", "").ToString().Trim());
                //seleccionamos la hoja numero i
                hoja.Activate();
                hojan.Activate();

                nom = libron.Worksheets[i].Name.ToString().Trim();
                ind = Array.IndexOf(nombre, libron.Worksheets[i].Name.ToString().Trim().ToUpper());

                //if (nom.ToUpper() == "CONTADOR") MessageBox.Show("gvfknjdkhjljkdfskl");

                //Barra de progreso
                if (this == null) return;
                Invoke(new System.Action(() => this.label1.Text = "Trabajando Hoja : [" + (Globals.ThisAddIn.Application.ActiveSheet).Name + "] .........."));
                if (this == null) return;
                Invoke(new System.Action(() => pgb_proceso.Value = pgb_proceso.Value + pgb_proceso.Step));
                
                //pasamos los datos
                EspacioFilas = 0;
                fila = 1;
                columna = 1;
                //Quitar proteccion de excel 
                //hojan.Unprotect(psw);
                hojan.Unprotect(ExcelAddIn.Access.Configuration.PwsExcel);

                do
                {
                    //Eliminar Comentarios
                    foreach (Comment a in ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Comments)
                    {
                        a.Delete();
                    }

                    columna = 1;
                    EspacioColumnas = 0;
                    do
                    {
                        if (ValidarString(hoja.Cells[fila, columna].Value).Trim().Length == 0)
                        {
                            EspacioColumnas++;
                        }
                        else
                        {
                            EspacioColumnas = 0;
                        }
                        columna++;
                    } while (EspacioColumnas < 10);//Si existen mas de 10 espacios en blanco ya no genera mas columnas

                    columna = 1;

                    if (ValidarString(hoja.Cells[fila, columna].Value).Trim().Length == 0 && ValidarString(hoja.Cells[fila, columna + 1].Value).Trim().Length == 0)
                    {
                        EspacioFilas++;
                    }
                    else
                    {
                        EspacioFilas = 0;
                    }
                    fila++;

                    if (EspacioFilas == 12 && ind != -1)
                    {
                        HojasSPR[ind, 1] = (fila - 12).ToString().Trim();
                    }
                } while (EspacioFilas < 12);//Si existen mas de 12 espacios en blanco ya no genera mas filas
            }
            //Ordenar Hojas Excel

            //deactivar mensajes alerta que genera al eliminar
            Globals.ThisAddIn.Application.DisplayAlerts = false;
            for (int j = 0; j < 5; j++)
            {
                for (int i = 1; i < numhojas; i++)
                {
                    try
                    {
                        nom = libron.Worksheets[i].Name.ToString().Trim();
                        ind = Array.IndexOf(nombre, libron.Worksheets[i].Name.ToString().Trim().ToUpper());
                        
                        //Barra de progreso
                        if (this == null) return;
                        Invoke(new System.Action(() => this.label1.Text = "Trabajando Hoja : [" + (Globals.ThisAddIn.Application.ActiveSheet).Name + "] .........."));
                        
                        if (ind != -1)
                        {
                            libron.Sheets[i].Activate();
                            
                            ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Move(Globals.ThisAddIn.Application.Worksheets[i]);
                            
                            cell1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[1, 1];
                            cell2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Cells[ValidarInt(ValidarString(HojasSPR[ind, 1]).Trim()), ValidarInt(ValidarString(HojasSPR[ind, 2]).Trim())];
                            range = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).get_Range(cell1, cell2);

                            range.Columns.AutoFit();
                        }
                        else
                        {
                            Excel.Worksheet m_objSheet = (Excel.Worksheet)(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.get_Item(nom));
                            m_objSheet.Visible = XlSheetVisibility.xlSheetHidden;
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }
            }
            

            for (int i = 1; i <= numhojas; i++)
            {
                try
                {
                    nom = libron.Worksheets[i].Name.ToString().Trim();
                    ind = Array.IndexOf(nombre, libron.Worksheets[i].Name.ToString().Trim().ToUpper());
                    if (ind == -1)
                    {
                        Excel.Worksheet m_objSheet = (Excel.Worksheet)(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.get_Item(nom));
                        m_objSheet.Visible = XlSheetVisibility.xlSheetHidden;
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }

            //activar mensajes alerta
            Globals.ThisAddIn.Application.DisplayAlerts = true;


            if (this == null) return;
            Invoke(new System.Action(() => this.pgb_proceso.Value = this.pgb_proceso.Maximum));
            Invoke(new System.Action(() => this.label1.Text = "Proceso Terminado."));
            
            Invoke(new System.Action(() => this.Visible = false));
            //this.Visible = false;
            Invoke(new System.Action(() => this.Close()));
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
        public String ValidarString(object val)
        {
            try { return Convert.ToString(val); }
            catch { return ""; }
        }
        public int ValidarInt(object val)
        {
            try { return Convert.ToInt32(val); }
            catch { return 0; }
        }
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
                //Color plomo de la etiqueta
                //if (ind == 15 || ind == 16 || ind == 19 || ind == 20 || ind == 21 || ind == 23 || ind == 24 || ind == 25)
                //{
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
                        
                        if (Vcon == "SI" || Vcon == "si")
                        {
                            res = false;
                            break;
                        }
                        if (Vcon == "NO" || Vcon == "no")
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
        private void frmCarga_FormClosed(object sender, FormClosedEventArgs e)
        {
            //if (this._ProcessJson == false)
            //{
            //    GuardarExcel();
            //    //Ribbon2 ai = new Ribbon2();

            //    //ai._Form = this;
            //    //ai.GuardarExcel();
            //}
        }
    }
}
