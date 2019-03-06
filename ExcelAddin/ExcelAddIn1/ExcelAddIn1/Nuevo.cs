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
using Newtonsoft.Json;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using ExcelAddIn.Access;

namespace ExcelAddIn1 {
    public partial class Nuevo : Base {
        public Nuevo() {
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
                    FillYears(cmbAnio);
                    FillTemplateType(cmbTipo);
                }
                else
                {
                    if (!_Connection)
                    {
                        MessageBox.Show(_Message.Replace("...", ", para crear los archivos base..."), "Creación de Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.btnCrear.Enabled = false;
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
            else{
                if (!Directory.Exists(_Path + "\\jsons")) {
                    Directory.CreateDirectory(_Path + "\\jsons");
                }
                if(!Directory.Exists(_Path + "\\templates")) {
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
        private void btnCancelar_Click(object sender, EventArgs e) {
            this.Close();
        }
        private void btnCrear_Click(object sender, EventArgs e) {
            string _Path = Configuration.Path;
            oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>($"{_Path}\\jsons\\Plantillas.json");
            int _IdTemplateType = (int)cmbTipo.SelectedValue, _Year = (int)cmbAnio.SelectedValue;
            
            if (_IdTemplateType == 0)
            {
                MessageBox.Show("Favor de seleccionar un Tipo de Plantilla.", "Tipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.cmbTipo.Focus();
                return;
            }
            if (_Year == 0)
            {
                MessageBox.Show("Favor de seleccionar el Año a Aplicar al Tipo de Plantilla.", "Año", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.cmbAnio.Focus();
                return;
            }
            
            oPlantilla _Template = _Templates.FirstOrDefault(o => o.IdTipoPlantilla == _IdTemplateType && o.Anio == _Year);

            if (_Template == null) {
                MessageBox.Show("No existe una plantilla para el tipo seleccionado, favor de seleccionar otro tipo o contactar al administrador.", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            
            fbdTemplate.ShowDialog();
            string _DestinationPath = fbdTemplate.SelectedPath;
            if(_DestinationPath == "") {
                MessageBox.Show("Debe especificar un ruta", "Ruta Invalida", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string _newTemplate = $"{_DestinationPath}\\{((oTipoPlantilla)cmbTipo.SelectedItem).Clave}-{cmbAnio.SelectedValue.ToString()}-{DateTime.Now.ToString("ddMMyyyyHHmmss")}.xlsm";
            GenerarArchivo(_Template, _newTemplate, ((oTipoPlantilla)cmbTipo.SelectedItem).Clave);
            this.Close();
        }
        protected void GenerarArchivo(oPlantilla _Template, string _DestinationPath, string _Tipo) {
            string _Path = Configuration.Path;
            oComprobacion[] _Comprobaciones = Assembler.LoadJson<oComprobacion[]>($"{_Path}\\jsons\\Comprobaciones.json");
            FileInfo _Excel = new FileInfo($"{_Path}\\templates\\{_Template.Nombre}");
            using(ExcelPackage _package = new ExcelPackage(_Excel)) {
                _package.Workbook.Worksheets.Add(_Tipo);
                foreach(oComprobacion _Comprobacion in _Comprobaciones.Where(o => o.IdTipoPlantilla == _Template.IdTipoPlantilla).ToArray()) {
                    ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Comprobacion.Destino.Anexo];
                    _Comprobacion.setFormulaExcel();
                    if(_Comprobacion.EsValida() && _Comprobacion.EsFormula())
                        _workSheet.Cells[_Comprobacion.Destino.CeldaExcel].Formula = _Comprobacion.FormulaExcel;
                    else if(_Comprobacion.EsValida() && !_Comprobacion.EsFormula())
                        _workSheet.Cells[_Comprobacion.Destino.CeldaExcel].Value = _Comprobacion.FormulaExcel;
                }
                //_package.Workbook.CreateVBAProject();
                byte[] _NewTemplate = _package.GetAsByteArray();
                File.WriteAllBytes(_DestinationPath, _NewTemplate);
            }
            Globals.ThisAddIn.Application.Visible = true;
            Globals.ThisAddIn.Application.Workbooks.Open(_DestinationPath);
            btnCancelar_Click(btnCancelar, null);
        }
    }
}