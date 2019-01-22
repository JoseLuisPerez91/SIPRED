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
using ExcelAddIn1.Assemblers;

namespace ExcelAddIn1 {
    public partial class Nuevo : Base {
        public Nuevo() {
            InitializeComponent();
            FillYears(cmbAnio);
            FillTemplateType(cmbTipo);
        }

        private void btnCancelar_Click(object sender, EventArgs e) {
            cmbAnio.SelectedIndex = 0;
            cmbTipo.SelectedIndex = 0;
        }

        private void btnCrear_Click(object sender, EventArgs e) {
            oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>($"{Environment.CurrentDirectory}\\jsons\\Templates.json");
            int _IdTemplateType = (int)cmbTipo.SelectedValue, _Year = (int)cmbAnio.SelectedValue;
            oPlantilla _Template = _Templates.FirstOrDefault(o => o.IdTipoPlantilla == _IdTemplateType && o.Anio == _Year);
            if(_Template != null) {
                MessageBox.Show("No existe una plantilla para el tipo seleccionado, favor de seleccionar otro tipo o contactar al administrador.", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            fbdTemplate.ShowDialog();
            sfdTemplate.OpenFile();
            if(fbdTemplate.SelectedPath == "") {
                MessageBox.Show("Debe especificar un ruta", "Ruta Invalida", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string _Path = fbdTemplate.SelectedPath;
            string _newTemplate = $"{_Path}\\{((oTipoPlantilla)cmbTipo.SelectedItem).Clave}-{cmbAnio.SelectedValue.ToString()}-{DateTime.Now.ToString("ddMMyyyyHHmmss")}.xlsm";
            string _currentTemplate = $"{Environment.CurrentDirectory}\\templates\\{_Template.Nombre}";
            Microsoft.Office.Interop.Excel.Application _current = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook _workbook = _current.Workbooks.Open(_currentTemplate, Microsoft.Office.Interop.Excel.XlUpdateLinks.xlUpdateLinksNever, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _workbook.SaveAs(_newTemplate, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _workbook.Close(false, Type.Missing, Type.Missing);
            _current.Quit();
            Globals.ThisAddIn.Application.Visible = true;
            Globals.ThisAddIn.Application.Workbooks.Open(_newTemplate);
            InicializarComprobaciones(_Path);
        }

        void InicializarComprobaciones(string _TemplateFile) {
            oComprobacion[] _Comprobaciones = Assembler.LoadJson<oComprobacion[]>($"{Environment.CurrentDirectory}\\jsons\\Comprobaciones.json");
            FileInfo _Excel = new FileInfo(_TemplateFile);
            using(ExcelPackage _package = new ExcelPackage(_Excel)) {
                foreach(oComprobacion _Comprobacion in _Comprobaciones) {
                    _Comprobacion.setCeldas();
                    ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Comprobacion.Destino.Anexo];
                    int _maxValue = _workSheet.Dimension.Rows + 1;
                    int _maxRow = (_workSheet.Dimension.Rows / 2) + (_workSheet.Dimension.Rows % 2);
                    for(int i = 1; i <= _maxRow; i++) {
                        _Comprobacion.Destino.Fila = (_workSheet.Cells[i, 1].Text == _Comprobacion.Destino.Indice) ? i : _Comprobacion.Destino.Fila;
                        _Comprobacion.Destino.Fila = (_workSheet.Cells[(_maxValue - i), 1].Text == _Comprobacion.Destino.Indice) ? _maxValue - i : _Comprobacion.Destino.Fila;
                        if(_Comprobacion.Destino.Fila > -1) {
                            oCelda[] _Celdas = _Comprobacion.Celdas.Where(o => o.Indice == _Comprobacion.Destino.Indice && o.Anexo == _Comprobacion.Destino.Anexo).ToArray();
                            oCelda[] _cCeldas = _Comprobacion.CeldasCondicion.Where(o => o.Indice == _Comprobacion.Destino.Indice && o.Anexo == _Comprobacion.Destino.Anexo).ToArray();
                            _Comprobacion.Destino.setCeldaExcel(_workSheet.Cells[_Comprobacion.Destino.Fila, _Comprobacion.Destino.Columna]);
                            foreach(oCelda _Celda in _Celdas) {
                                _Celda.Fila = _Comprobacion.Destino.Fila;
                                _Celda.setCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna]);
                            }
                            foreach(oCelda _Celda in _cCeldas) {
                                _Celda.Fila = _Comprobacion.Destino.Fila;
                                _Celda.setCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna]);
                            }
                            oCelda[] _Faltantes = _Comprobacion.Celdas.Where(o => o.Fila == -1).ToArray();
                            foreach(oCelda _Faltante in _Faltantes) {
                                oCelda _Result = _Comprobaciones.Where(o => o.Destino != null && o.Destino.Indice == _Faltante.Indice && o.Destino.Anexo == _Faltante.Anexo.ToUpper()).Select(o => o.Destino).FirstOrDefault();
                                if(_Result != null) {
                                    _Faltante.Fila = _Result.Fila;
                                    _Faltante.setCeldaExcel(_workSheet.Cells[_Faltante.Fila, _Faltante.Columna]);
                                }
                                if(_Result == null) {
                                    ExcelWorksheet _ws = _package.Workbook.Worksheets[_Faltante.Anexo];
                                    int _mv = _ws.Dimension.Rows + 1;
                                    int _mr = (_ws.Dimension.Rows / 2) + (_ws.Dimension.Rows % 2);
                                    for(int j = 1; j <= _mr; j++) {
                                        _Faltante.Fila = (_ws.Cells[j, 1].Text == _Faltante.Indice) ? j : _Faltante.Fila;
                                        _Faltante.Fila = (_ws.Cells[(_mv - j), 1].Text == _Faltante.Indice) ? _mv - j : _Faltante.Fila;
                                        if(_Faltante.Fila > -1) {
                                            _Faltante.setCeldaExcel(_ws.Cells[_Faltante.Fila, _Faltante.Columna]);
                                            break;
                                        }
                                    }
                                }
                            }
                            oCelda[] _cFaltantes = _Comprobacion.CeldasCondicion.Where(o => o.Fila == -1).ToArray();
                            foreach(oCelda _Faltante in _cFaltantes) {
                                oCelda _Result = _Comprobaciones.Where(o => o.Destino != null && o.Destino.Indice == _Faltante.Indice && o.Destino.Anexo == _Faltante.Anexo.ToUpper()).Select(o => o.Destino).FirstOrDefault();
                                if(_Result != null) {
                                    _Faltante.Fila = _Result.Fila;
                                    _Faltante.setCeldaExcel(_workSheet.Cells[_Faltante.Fila, _Faltante.Columna]);
                                }
                                if(_Result == null) {
                                    ExcelWorksheet _ws = _package.Workbook.Worksheets[_Faltante.Anexo];
                                    int _mv = _ws.Dimension.Rows + 1;
                                    int _mr = (_ws.Dimension.Rows / 2) + (_ws.Dimension.Rows % 2);
                                    for(int j = 1; j <= _mr; j++) {
                                        _Faltante.Fila = (_ws.Cells[j, 1].Text == _Faltante.Indice) ? j : _Faltante.Fila;
                                        _Faltante.Fila = (_ws.Cells[(_mv - j), 1].Text == _Faltante.Indice) ? _mv - j : _Faltante.Fila;
                                        if(_Faltante.Fila > -1) {
                                            _Faltante.setCeldaExcel(_ws.Cells[_Faltante.Fila, _Faltante.Columna]);
                                            break;
                                        }
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
        }
    }
}