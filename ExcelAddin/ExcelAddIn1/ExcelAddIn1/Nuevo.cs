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
//using ExcelAddIn1.Assemblers;

namespace ExcelAddIn1 {
    public partial class Nuevo : Base {
        public Nuevo() {
            string _Path = ExcelAddIn.Access.Configuration.Path;
            InitializeComponent();

            Cursor = System.Windows.Forms.Cursors.WaitCursor;
            if (Directory.Exists(_Path + "\\jsons") && Directory.Exists(_Path + "\\templates"))
            {
                if (File.Exists(_Path + "\\jsons\\TiposPlantillas.json")) {
                    FillYears(cmbAnio);
                    FillTemplateType(cmbTipo);
                }
                else
                {
                    MessageBox.Show("Los archivos base serán generados... Click en el botón Aceptar para continuar.", "Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    KeyValuePair<bool, string[]> _result = new lSerializados().ObtenerSerializados();

                    FillYears(cmbAnio);
                    FillTemplateType(cmbTipo);
                }
            }
            else{
                MessageBox.Show("Los archivos base serán generados... Click en el botón Aceptar para continuar.", "Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (!Directory.Exists(_Path + "\\jsons")) {
                    Directory.CreateDirectory(_Path + "\\jsons");
                }
                if(!Directory.Exists(_Path + "\\templates")) {
                    Directory.CreateDirectory(_Path + "\\templates");
                }

                KeyValuePair<bool, string[]> _result = new lSerializados().ObtenerSerializados();
                FillYears(cmbAnio);
                FillTemplateType(cmbTipo);
            }
            Cursor = System.Windows.Forms.Cursors.Default;
        }

        private void btnCancelar_Click(object sender, EventArgs e) {
            cmbAnio.SelectedIndex = 0;
            cmbTipo.SelectedIndex = 0;
            cmbTipo.Focus();
        }

        private void btnCrear_Click(object sender, EventArgs e) {
            string _Path = ExcelAddIn.Access.Configuration.Path;
            oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>($"{_Path}\\jsons\\Plantillas.json");
            int _IdTemplateType = (int)cmbTipo.SelectedValue, _Year = (int)cmbAnio.SelectedValue;
            oPlantilla _Template = _Templates.FirstOrDefault(o => o.IdTipoPlantilla == _IdTemplateType && o.Anio == _Year);
            if(_Template == null) {
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
            string _Path = ExcelAddIn.Access.Configuration.Path;
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
                _package.Workbook.CreateVBAProject();
                byte[] _NewTemplate = _package.GetAsByteArray();
                File.WriteAllBytes(_DestinationPath, _NewTemplate);
            }
            Globals.ThisAddIn.Application.Visible = true;
            Globals.ThisAddIn.Application.Workbooks.Open(_DestinationPath);
            btnCancelar_Click(btnCancelar, null);
        }
    }
}