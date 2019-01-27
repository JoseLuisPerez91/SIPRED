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

namespace ExcelAddIn1 {
    public partial class Cruce : Base {
        public Cruce() {
            InitializeComponent();
        }

        private void btnAceptar_Click(object sender, EventArgs e) {
            string _Path = ExcelAddIn.Access.Configuration.Path;
            oTipoPlantilla[] _TemplateTypes = Assembler.LoadJson<oTipoPlantilla[]>($"{_Path}\\jsons\\TiposPlantillas.json");
            oCruce[] _Cruces = Assembler.LoadJson<oCruce[]>($"{_Path}\\jsons\\Cruces.json");
            FileInfo _Excel = new FileInfo(Globals.ThisAddIn.Application.ThisWorkbook.FullName);
            oTipoPlantilla _TemplateType = null;
            using(ExcelPackage _package = new ExcelPackage(_Excel)) {
                foreach(oTipoPlantilla _TT in _TemplateTypes) {
                    if(_package.Workbook.Worksheets.Where(o => o.Name == _TT.Clave).FirstOrDefault() != null)
                        _TemplateType = _TT;
                }
                if(_TemplateType != null) {
                    _package.Workbook.Worksheets.Add("Test");
                    ExcelWorksheet _wsTest = _package.Workbook.Worksheets.First(o => o.Name == "Test");
                    foreach(oCruce _Cruce in _Cruces.Where(o => o.IdTipoPlantilla == _TemplateType.IdTipoPlantilla)) {
                        _Cruce.setCeldas();
                        foreach(oCelda _Celda in _Cruce.CeldasFormula) {
                            ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Celda.Anexo];
                            int _maxValue = _workSheet.Dimension.Rows + 1;
                            int _maxRow = (_workSheet.Dimension.Rows / 2) + (_workSheet.Dimension.Rows % 2);
                            for(int i = 1; i <= _maxRow; i++) {
                                _Celda.Fila = (_workSheet.Cells[i, 1].Text == _Celda.Indice) ? i : _Celda.Fila;
                                _Celda.Fila = (_workSheet.Cells[(_maxValue - i), 1].Text == _Celda.Indice) ? _maxValue - i : _Celda.Fila;
                                if(_Celda.Fila > -1) _Celda.setCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna], _Celda.Anexo);
                            }
                        }
                        foreach(oCelda _Celda in _Cruce.CeldasCondicion) {
                            ExcelWorksheet _workSheet = _package.Workbook.Worksheets[_Celda.Anexo];
                            if(_workSheet != null) {
                                int _maxValue = _workSheet.Dimension.Rows + 1;
                                int _maxRow = (_workSheet.Dimension.Rows / 2) + (_workSheet.Dimension.Rows % 2);
                                for(int i = 1; i <= _maxRow; i++) {
                                    _Celda.Fila = (_workSheet.Cells[i, 1].Text == _Celda.Indice) ? i : _Celda.Fila;
                                    _Celda.Fila = (_workSheet.Cells[(_maxValue - i), 1].Text == _Celda.Indice) ? _maxValue - i : _Celda.Fila;
                                    if(_Celda.Fila > -1) _Celda.setCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna], _Celda.Anexo);
                                }
                            }
                        }
                        _Cruce.setFormulaExcel();
                        _wsTest.Cells["A1"].Formula = _Cruce.FormulaExcel;
                        _wsTest.Cells["A1"].Calculate();
                        if(_Cruce.CondicionExcel != "") {
                            _wsTest.Cells["A2"].Formula = _Cruce.CondicionExcel;
                            _wsTest.Cells["A2"].Calculate();
                            _Cruce.ResultadoCondicion = _wsTest.Cells["A2"].Value.ToString();
                        }
                        _Cruce.ResultadoFormula = _wsTest.Cells["A1"].Value.ToString();
                    }
                } else if(_TemplateType == null)
                    MessageBox.Show("Archivo no valido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}