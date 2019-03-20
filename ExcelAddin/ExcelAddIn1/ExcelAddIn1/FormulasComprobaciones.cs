using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Newtonsoft.Json;
using ExcelAddIn.Access;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using System.Net;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Core;
using Microsoft.Win32;

namespace ExcelAddIn1
{
    public partial class FormulasComprobaciones : Base
    {
        public Form _Form;
        public oPlantilla _Template;
        public string _Tipo;
        public bool _formulas;
        public string _Origen;
        public bool _Open;
        public FormulasComprobaciones()
        {
            string _Path = Configuration.Path;
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            InitializeComponent();

            if (Directory.Exists(_Path + "\\jsons") && Directory.Exists(_Path + "\\templates"))
            {
                if (File.Exists(_Path + "\\jsons\\Comprobaciones.json"))
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
                        this.btnGenerar.Enabled = false;
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
                _FileJsonfrm._window = this.Text;
                _FileJsonfrm.Show();
                return;
            }
        }
        private void btnGenerar_Click(object sender, EventArgs e)
        {
            //Variables generales.
            string _Path = Configuration.Path;
            int x = 0;
            double r = 0;
            int progress = 0;
            oComprobacion[] _Comprobaciones = Assembler.LoadJson<oComprobacion[]>($"{_Path}\\jsons\\Comprobaciones.json");
            //Libro Actual de Excel.
            Excel.Worksheet xlSht;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            string _Name = Globals.ThisAddIn.Application.ActiveWorkbook.Name;
            string[] _aName = _Name.Split('-');
            string _anio = _aName[1];
            string _IdTipo = "";
            string _TipoFile = _aName[0].ToString();
            string _DestinationPath = "";
            string _newTemplate = "";

            _Name = _aName[2].ToString();
            _IdTipo = _Name.Split('_')[1].ToString();

            //Cuándo es para transferir, pide la ruta en donde guardar el archivo a transferir.
            if (!_formulas)
            {
                for (int y = 0; y < 1;)
                {
                    fbdTemplate.ShowDialog();
                    _DestinationPath = fbdTemplate.SelectedPath;
                    y = 1;
                    if (_DestinationPath == "")
                    {
                        MessageBox.Show("Debe especificar un ruta", "Ruta Invalida", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        y = 0;
                    }
                }

                // el nombre de una Key debe incluir un root valido.
                const string userRoot = "HKEY_CURRENT_USER";
                const string subkey = "Software\\Microsoft\\Office\\Excel\\Addins\\SAT.Dictamenes.SIPRED.Client";
                const string keyName = userRoot + "\\" + subkey;
                object addInName = "SAT.Dictamenes.SIPRED.Client";

                Registry.SetValue(keyName, "LoadBehavior", 3);
                Globals.ThisAddIn.Application.COMAddIns.Item(ref addInName).Connect = true;
                string _sOpen = wb.Worksheets.Item[wb.Worksheets.Count].Name;

                _newTemplate = $"{_DestinationPath}\\Transferencia-{_TipoFile}-{_anio}-{DateTime.Now.ToString("ddMMyyyyHHmmss")}_{_IdTipo}_{_anio}.xlsm";

                for(int _wCount = 1; _wCount <= wb.Worksheets.Count; _wCount++)
                {
                    if (wb.Worksheets.Item[_wCount].Name == "SIPRED")
                    {
                        wb.Worksheets.Item[_wCount].Delete();
                    }
                    else
                    {
                        if (wb.Worksheets.Item[_wCount].Name == "_Open")
                        {
                            wb.Worksheets.Item[_wCount].Delete();
                        }
                    }
                }

                wb.SaveCopyAs(_newTemplate);
                wb.Save();
                wb.Close();

                Globals.ThisAddIn.Application.Visible = true;
                Globals.ThisAddIn.Application.Workbooks.Open(_newTemplate);

                wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                x = 0;
                foreach (oComprobacion _Comprobacion in _Comprobaciones.Where(o => o.IdTipoPlantilla == Convert.ToInt32(_IdTipo)).ToArray())
                {
                    _Comprobacion.setFormulaExcel();
                    xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_Comprobacion.Destino.Anexo);
                    Excel.Range _Range = (Excel.Range)xlSht.get_Range(_Comprobacion.Destino.CeldaExcel);

                    object _valor = _Range.Value;
                    //_Range.NumberFormat = "0.00";
                    if (_Comprobacion.EsValida() && _Comprobacion.EsFormula())
                    {
                        try
                        {
                            if (!_Open && _sOpen == "SIPRED")
                            {
                                string _Columna1 = Generales.ColumnAdress(_Range.Column);
                                string _Renglon1 = (_Range.Row - 1).ToString();
                                object _Celda1 = _Columna1 + "" + _Renglon1;
                                Excel.Range _Range1 = (Excel.Range)xlSht.get_Range(_Celda1);
                                object _valor1 = _Range1.Value;

                                //Papá de los Indices.
                                if (_Range1.HasFormula)
                                {
                                    _Range1.Formula = "";
                                    _Range1.Value = Convert.ToDecimal(_valor1.ToString());
                                }
                                //Indices
                                if (!_Range.HasFormula)
                                {
                                    int _Rango = 1;
                                    string _Renglon;
                                    string _Columna;
                                    object _Celda;
                                    for (int xx = 0; xx < 1;)
                                    {
                                        _Columna = Generales.ColumnAdress(_Range.Column);
                                        _Renglon = (_Range.Row + _Rango).ToString();
                                        _Celda = _Columna + "" + _Renglon;
                                        Excel.Range _Range2 = (Excel.Range)xlSht.get_Range(_Celda);

                                        if (!_Range2.HasFormula)
                                        {
                                            xx = 0;
                                            _Rango += 1;
                                        }
                                        else
                                        {
                                            _Range.Formula = "";
                                            _Range.Value = Convert.ToDecimal(_valor.ToString());
                                            //_Range2.FormulaHidden = true;
                                            xx = 1;
                                            //_Rango = 1;
                                        }
                                    }
                                }
                                else
                                {
                                    _Range.Formula = "";
                                    _Range.Value = Convert.ToDecimal(_valor.ToString());
                                    //_Range.FormulaHidden = true;
                                }
                            }
                        }
                        catch { }
                    }
                }
                wb.Save();
            }
            //Asigna valores vacios a las celdas de las formulas y de tipo "General".
            if (_formulas)
            {
                x = 0;
                foreach (oComprobacion _Comprobacion in _Comprobaciones.Where(o => o.IdTipoPlantilla == Convert.ToInt32(_IdTipo)).ToArray())
                {
                    _Comprobacion.setFormulaExcel();
                    xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_Comprobacion.Destino.Anexo);

                    string _fExcel = _Comprobacion.FormulaExcel.Replace("SUM", "").Replace("(", "").Replace(")", "").Replace("+0", "").Replace("*", "+").Replace("/", "+").Replace("IF", "").Replace("<0", "").Replace(">0", "+").Replace(",0)", "").Replace(",", "+").Replace("-", "+").Replace(">", "+").Replace("<", "+").Replace("=", "+");
                    string[] _sfExcel = _fExcel.Split('+');

                    for (int z = 0; z < _sfExcel.Length; z++)
                    {
                        if (_sfExcel[z] != "")
                        {
                            decimal temp = 0;
                            if (!decimal.TryParse(_sfExcel[z], out temp))
                            {
                                Excel.Range _Celda = (Excel.Range)xlSht.get_Range(_sfExcel[z]);
                                _Celda.NumberFormat = "0.00";
                                if (!_Open)
                                {
                                    _Celda.Value = "";
                                }
                            }
                        }
                    }
                    //Barra de Progreso.
                    x++;
                    r = x % 16;
                    if (r == 0.00)
                    {
                        progress += 10;
                        if (progress < 100)
                        {
                            fnProgressBar(progress);
                        }
                    }
                }
                x = 0;
                //Asigna las formulas a las celdas al crear un nuevo archivo
                //De lo contrario si es transferir quita las formulas y asigna el valor del resultado de la formula.
                //Se agina el progreso del ProgessBar según la cantidad de celdas divididas entre 16.
                foreach (oComprobacion _Comprobacion in _Comprobaciones.Where(o => o.IdTipoPlantilla == Convert.ToInt32(_IdTipo)).ToArray())
                {
                    _Comprobacion.setFormulaExcel();
                    xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_Comprobacion.Destino.Anexo);
                    Excel.Range _Range = (Excel.Range)xlSht.get_Range(_Comprobacion.Destino.CeldaExcel);

                    if (x == 0)
                    {
                        xlSht.Activate();
                    }
                    _Range.NumberFormat = "0.00";
                    if (_Comprobacion.EsValida() && _Comprobacion.EsFormula())
                    {
                        _Range.Formula = $"={_Comprobacion.FormulaExcel}";
                    }
                    else if (_Comprobacion.EsValida() && !_Comprobacion.EsFormula())
                    {
                        _Range.Value = _Comprobacion.FormulaExcel;
                    }
                    //Barra de Progreso.
                    x++;
                    r = x % 16;
                    if (r == 0.00)
                    {
                        progress += 10;
                        if (progress < 100)
                        {
                            fnProgressBar(progress);
                        }
                    }
                }
                //Se guarda el archivo original.
                wb.Save();
            }
            if (_Form != null)
            {
                _Form.Close();
            }
            this.Close();
        }
        private void FormulasComprobaciones_Load(object sender, EventArgs e)
        {
            string _Message = "";
            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);

            if (_Excel.Extension != ".xlsm")
            {
                MessageBox.Show("Archivo no válido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;
            }

            if (!_formulas)
            { 
                _Message = "Clic en [Aceptar] para Transfirir el Archivo... Espere mientras termina el proceso.";
                this.btnGenerar.Visible = true;
                this.btnGenerar.Enabled = true;
                this.Height = 122;
                Invoke(new System.Action(() => this.label1.Text = _Message));
            }
        }
        private void FormulasComprobaciones_Shown(object sender, EventArgs e)
        {
            string _Message = "";
            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);

            if (_Excel.Extension != ".xlsm")
            {
                MessageBox.Show("Archivo no válido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
                return;
            }
            if (_formulas)
            {
                _Message = "Generando las formulas de Comprobaciones... Espere mientras termina el proceso.";
                this.btnGenerar.Visible = false;
                this.btnGenerar.Enabled = false;
                this.Height = 97;
                Invoke(new System.Action(() => this.label1.Text = _Message));
                btnGenerar_Click(sender, e);
            }
        }
        private void fnProgressBar(int _Progress)
        {
            Invoke(new System.Action(() => this.pgbFile.Value = _Progress));
        }
    }
}
