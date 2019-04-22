using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAddIn.Access;
using ExcelAddIn.Logic;
using ExcelAddIn.Objects;
using Microsoft.Win32;
using System.IO;
using Microsoft.Office;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace ExcelAddIn1
{
    public partial class ConversionMasiva : Base
    {
        private DataTable dtAnexo = new DataTable();
        private bool _Error;
        private string _sError;
        public ConversionMasiva()
        {
            InitializeComponent();
            //Creación de la tabla de Anexos.
            dtAnexo.Columns.Add(new DataColumn("Anexo", typeof(string)));
            dtAnexo.Columns.Add(new DataColumn("Descripcion", typeof(string)));
            dtAnexo.Columns.Add(new DataColumn("Convertir", typeof(bool)));
            dtAnexo.TableName = "Anexos";
        }

        private void btnConvertir_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            int _Year = (int)cmbAnio.SelectedValue;

            if (_Year == 0)
            {
                MessageBox.Show("Favor de seleccionar el Año a Aplicar al Tipo de Plantilla.", "Año", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                this.cmbAnio.Focus();
                return;
            }
            if (this.txtRuta.Text == "")
            {
                MessageBox.Show("Debe especificar un ruta", "Ruta Invalida", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            Excel.Workbook _wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            _wb.Saved = false;
            
            string _Path = Configuration.Path;
            int _Rows = dgvAnexos.RowCount;
            oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>($"{_Path}\\jsons\\Plantillas.json");
            string _DestinationPath = "";
            int _Avance = 0;
            //Libro Actual de Excel.
            List<Excel.Worksheet> _xAnexos = new List<Excel.Worksheet>();
            List<string> _sAnexo = new List<string>();
            Excel.Worksheet xlSht;

            //Barra de Progreso.
            fnProgressBar(5);
            for (int x = 0; x <= _Rows-1; x++)
            {
                Cursor.Current = Cursors.WaitCursor;
                if (Convert.ToBoolean(dgvAnexos.Rows[x].Cells["Convertir"].Value))
                {
                    _xAnexos.Add((Excel.Worksheet)_wb.Worksheets.get_Item(dgvAnexos.Rows[x].Cells["Anexo"].Value));
                    xlSht = (Excel.Worksheet)_wb.Worksheets.get_Item(dgvAnexos.Rows[x].Cells["Anexo"].Value);
                    _sAnexo.Add(dgvAnexos.Rows[x].Cells["Anexo"].Value.ToString());
                }
            }

            //this.TopMost = false;
            //// el nombre de una Key debe incluir un root valido.
            //const string userRoot = "HKEY_CURRENT_USER";
            //const string subkey = "Software\\Microsoft\\Office\\Excel\\Addins\\SAT.Dictamenes.SIPRED.Client";
            //const string keyName = userRoot + "\\" + subkey;

            //object addInName = "SAT.Dictamenes.SIPRED.Client";
            //Registry.SetValue(keyName, "LoadBehavior", 0);
            //Globals.ThisAddIn.Application.COMAddIns.Item(ref addInName).Connect = false;
            //Globals.ThisAddIn.Application.Visible = true;

            _wb.Save();
            //this.TopMost = true;
            //this.Focus();
            oPlantilla _Template = _Templates.FirstOrDefault(o => o.IdTipoPlantilla == 1 && o.Anio == _Year);
            _DestinationPath = this.txtRuta.Text + $"\\{this.txtRFC.Text}-{_Year}-{DateTime.Now.ToString("ddMMyyyy")}_{_Template.IdTipoPlantilla}_{_Year}.xlsm";

            if (File.Exists(_DestinationPath))
            {
                File.Delete(_DestinationPath);
            }
            //Barra de Progreso.
            fnProgressBar(8);
            File.Copy($"{_Path}\\templates\\{_Template.Nombre}", _DestinationPath);
            Globals.ThisAddIn.Application.Workbooks.Open(_DestinationPath);

            //Libro Actual de Excel.
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            int count = wb.Worksheets.Count;
            bool _Sipred = false;

            for (int y = 1; y <= count; y++)
            {
                Cursor.Current = Cursors.WaitCursor;
                xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(y);

                if (xlSht.Name == "SIPRED")
                {
                    _Sipred = true;
                }
            }
            //Barra de Progreso.
            fnProgressBar(10);

            if (!_Sipred)
            {
                xlSht = (Excel.Worksheet)wb.Worksheets.Add(Type.Missing, wb.Worksheets[count], Type.Missing, Type.Missing);
                xlSht.Name = "SIPRED";
                wb.Save();
            }

            //Barra de Progreso.
            fnProgressBar(11);
            _Avance = 11;
            oComprobacion[] _Comprobaciones = Assembler.LoadJson<oComprobacion[]>($"{_Path}\\jsons\\Comprobaciones.json");
            foreach (oComprobacion _Comprobacion in _Comprobaciones.Where(o => o.IdTipoPlantilla == Convert.ToInt32(1)).ToArray())
            {
                Cursor.Current = Cursors.WaitCursor;
                _Avance += 1;
                if(_Avance<100)
                {
                    fnProgressBar(_Avance);
                }
                else
                {
                    fnProgressBar(99);
                    _Avance = 11;
                }
                
                xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_Comprobacion.Destino.Anexo);
                xlSht.Unprotect(ExcelAddIn.Access.Configuration.PwsExcel);
            }
            
            //Barra de Progreso.
            fnProgressBar(34);
            _Avance = 34;
            foreach (oComprobacion _Comprobacion in _Comprobaciones.Where(o => o.IdTipoPlantilla == Convert.ToInt32(1)).ToArray())
            {
                Cursor.Current = Cursors.WaitCursor;
                _Avance += 1;
                if (_Avance < 100)
                {
                    fnProgressBar(_Avance);
                }
                else
                {
                    fnProgressBar(99);
                    _Avance = 34;
                }

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
                            _Celda.Value = "";
                        }
                    }
                }
            }

            _Avance = 57;
            //Barra de Progreso.
            fnProgressBar(57);
            //Ciclo para copiar los valores del XSPR al nuevo Template
            for (int x = 0; x <= 7; x++)
            {
                Cursor.Current = Cursors.WaitCursor;
                _Avance += 2;
                switch (x)
                {
                    case 0:
                        if (Convert.ToBoolean(dtAnexo.Rows[x]["Convertir"]) == true)
                        {
                            TransferConvertionValue(wb, _xAnexos[x], _sAnexo[x], 12, 1000);
                        }
                        fnProgressBar(_Avance);
                        break;
                    case 1:
                        if (Convert.ToBoolean(dtAnexo.Rows[x]["Convertir"]) == true)
                        {
                            TransferAnexo2(wb, _xAnexos[x], _sAnexo[x], 3, 1000);
                            //TransferConvertionValue(wb, _xAnexos[x], 3, 1000);
                        }
                        fnProgressBar(_Avance);
                        break;
                    case 2:
                        if (Convert.ToBoolean(dtAnexo.Rows[x]["Convertir"]) == true)
                        {
                            TransferAnexo3(wb, _xAnexos[x], _sAnexo[x], 3, 1000);
                        }
                        fnProgressBar(_Avance);
                        break;
                    case 3:
                        if (Convert.ToBoolean(dtAnexo.Rows[x]["Convertir"]) == true)
                        {
                            TransferConvertionValue(wb, _xAnexos[x], _sAnexo[x], 10, 1000);
                        }
                        fnProgressBar(_Avance);
                        break;
                    case 4:
                        if (Convert.ToBoolean(dtAnexo.Rows[x]["Convertir"]) == true)
                        {
                            TransferConvertionValue(wb, _xAnexos[x], _sAnexo[x], 6, 1000);
                        }
                        fnProgressBar(_Avance);
                        break;
                    case 5:
                        if (Convert.ToBoolean(dtAnexo.Rows[x]["Convertir"]) == true)
                        {
                            TransferConvertionValue(wb, _xAnexos[x], _sAnexo[x], 4, 1000);
                        }
                        fnProgressBar(_Avance);
                        break;
                    case 6:
                        if (Convert.ToBoolean(dtAnexo.Rows[x]["Convertir"]) == true)
                        {
                            TransferConvertionValue(wb, _xAnexos[x], _sAnexo[x], 8, 1000);
                        }
                        fnProgressBar(_Avance);
                        break;
                    case 7:
                        if (Convert.ToBoolean(dtAnexo.Rows[x]["Convertir"]) == true)
                        {
                            TransferAnexo8(wb, _xAnexos[x], _sAnexo[x], 3, 1000);
                        }
                        fnProgressBar(_Avance);
                        break;
                }
            }

            _Avance = 73;
            //Barra de Progreso.
            fnProgressBar(73);
            //Proceso Json para asignar formulas.
            foreach (oComprobacion _Comprobacion in _Comprobaciones.Where(o => o.IdTipoPlantilla == Convert.ToInt32(1)).ToArray())
            {
                Cursor.Current = Cursors.WaitCursor;
                _Avance += 1;
                if (_Avance < 100)
                {
                    fnProgressBar(_Avance);
                }
                else
                {
                    fnProgressBar(99);
                    _Avance = 73;
                }

                string _Formula = "";
                string[] _aFormula = _Comprobacion.Formula.Split('=');
                string _sFormula = Convert.ToString(_aFormula[1]);
                string _oFormula = _sFormula;
                string[] _aAnexo = _Comprobacion.Destino.Anexo.Split(' ');
                string _nAnexo = "0" + Convert.ToString(_aAnexo[1]);
                string[] _Indices;

                if (_nAnexo.Length == 3)
                {
                    _nAnexo = _nAnexo.Substring(1, 2);
                }

                _Formula = _sFormula.Replace($"[{_nAnexo},", "");
                _Formula = _Formula.Replace("SUM", "");
                _Formula = _Formula.Replace("(", "");
                _Formula = _Formula.Replace(")", "");
                _Formula = _Formula.Replace(":", "+");
                _Formula = _Formula.Replace("-", "+");
                _Formula = _Formula.Replace("*", "+");
                _Formula = _Formula.Replace("]", "");
                _Formula = _Formula.Replace("/", "+");
                _Formula = _Formula.Replace(" ", "");
                _Indices = _Formula.Split('+');

                xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_Comprobacion.Destino.Anexo);

                for (int arr = 0; arr < _Indices.Length; arr++)
                {
                    if (_Indices[arr] != "")
                    {
                        int _out;
                        if(!Int32.TryParse(_Indices[arr], out _out))
                        {
                            string sAnexo = _Comprobacion.Destino.Anexo;
                            string[] _Valores = _Indices[arr].Split(',');

                            for (int a = 1; a < 1000; a++)
                            {
                                Excel.Range _CeldaA = (Excel.Range)xlSht.get_Range("A" + a.ToString());
                                if (Convert.ToString(_CeldaA.Value) == Convert.ToString(_Valores[0]))
                                {
                                    string _Rango = Generales.ColumnAdress(Convert.ToInt32(_Valores[1])) + a.ToString();
                                    Int32 _iColumna = Convert.ToInt32(_Valores[1]);

                                    _CeldaA = (Excel.Range)xlSht.get_Range(_Rango);
                                    if(_iColumna>2 && _iColumna<5)
                                    {
                                        _CeldaA.NumberFormat = "0";
                                    }
                                    else
                                    {
                                        _CeldaA.NumberFormat = "0.00";
                                    }
                                    
                                    _oFormula = _oFormula.Replace($"[{_nAnexo},", "");
                                    _oFormula = _oFormula.Replace($"{_Valores[0]},{_Valores[1]}", _Rango);
                                    _oFormula = _oFormula.Replace("]", "");
                                    a = 1001;
                                }
                            }
                        }
                    }
                }
                for (int a = 1; a < 1000; a++)
                {
                    Excel.Range _Celda = (Excel.Range)xlSht.get_Range("A" + a.ToString());
                    if (Convert.ToString(_Celda.Value) == Convert.ToString(_Comprobacion.Destino.Indice))
                    {
                        string _Rango = Generales.ColumnAdress(Convert.ToInt32(_Comprobacion.Destino.Columna)) + a.ToString();

                        _Celda = (Excel.Range)xlSht.get_Range(_Rango);
                        if (_Comprobacion.Destino.Columna > 2 && _Comprobacion.Destino.Columna < 5)
                        {
                            _Celda.NumberFormat = "0";
                        }
                        else
                        {
                            _Celda.NumberFormat = "0.00";
                        }
                        _Celda.Formula = $"={_oFormula}";
                        a = 1001;
                    }
                }

                if (_Comprobacion.Destino.Anexo == "ANEXO 1")
                {
                    xlSht.Activate();
                }
            }

            fnProgressBar(99);
            wb.Save();
            fnProgressBar(100);
            Cursor.Current = Cursors.Default;
            _wb.Close();
            this.Close();
        }

        private void TransferConvertionValue(Excel.Workbook wb, Excel.Worksheet _xAnexo, string _sAnexo, int _inicial, int _final)
        {
            Excel.Range _Range;
            Excel.Range _RangeP;
            Excel.Range _RangeN;
            Excel.Range _RangeS;
            string _sValor;
            Int64 _Valor;
            int _Hijo;
            Int64 _Resultado;
            int _Vacio = 0, _IndiceI = 0, _IndiceF = 0;
            Excel.Worksheet xlSht = _xAnexo;
            Excel.Worksheet _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);

            for (int y = _inicial; y <= _final; y++)
            {
                _RangeP = (Excel.Range)xlSht.get_Range($"A{(y - 1).ToString()}");
                _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                _RangeS = (Excel.Range)xlSht.get_Range($"A{(y + 1).ToString()}");

                if (_Range.Value == "")
                {
                    _Vacio += 1;
                }
                if (_Range.Value != "")
                {
                    _Vacio = 0;
                    if (Int64.TryParse(_Range.Value, out _Valor))
                    {
                        _Valor = Convert.ToInt64(_Range.Value);
                        _sValor = Convert.ToString(_Range.Value);
                        _Hijo = Convert.ToInt32(_sValor.Substring(_sValor.Length - 4, 4));

                        if (_Hijo >= 100)
                        {
                            if(_IndiceF == 0)
                            { _IndiceI = y; }
                            _IndiceF += 1;
                            
                            _RangeN = _xlSht.get_Range(string.Format("{0}:{0}", (y), Type.Missing));
                            _RangeN.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                            //Columna A
                            _RangeN = (Excel.Range)_xlSht.get_Range($"A{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                            _RangeN.Value = _Range.Value;
                            Generales.AddNamedRange(y, 1, "IA_0" + _RangeN.Value, _xlSht);
                            //Columna B
                            _RangeN = (Excel.Range)_xlSht.get_Range($"B{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"B{y.ToString()}");
                            _RangeN.Value = _Range.Value;
                            //Columna D
                            _RangeN = (Excel.Range)_xlSht.get_Range($"D{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"C{y.ToString()}");
                            string _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                                
                            }
                            //ANEXO 7
                            if(_sAnexo == "ANEXO 7")
                            {
                                //Columna K
                                _RangeN = (Excel.Range)_xlSht.get_Range($"K{(y).ToString()}");
                                _Range = (Excel.Range)xlSht.get_Range($"J{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna R
                                _RangeN = (Excel.Range)_xlSht.get_Range($"R{(y).ToString()}");
                                _Range = (Excel.Range)xlSht.get_Range($"Q{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna Y
                                _RangeN = (Excel.Range)_xlSht.get_Range($"Y{(y).ToString()}");
                                _Range = (Excel.Range)xlSht.get_Range($"X{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna AF
                                _RangeN = (Excel.Range)_xlSht.get_Range($"AF{(y).ToString()}");
                                _Range = (Excel.Range)xlSht.get_Range($"AE{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna AH
                                _RangeN = (Excel.Range)_xlSht.get_Range($"AH{(y).ToString()}");
                                _Range = (Excel.Range)xlSht.get_Range($"AG{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna AJ
                                _RangeN = (Excel.Range)_xlSht.get_Range($"AJ{(y).ToString()}");
                                _Range = (Excel.Range)xlSht.get_Range($"AI{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if(_IndiceF > 0)
                            {
                                //Columna C
                                _RangeN = (Excel.Range)_xlSht.get_Range($"C{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0";
                                _RangeN.Formula = $"=SUM(C{_IndiceI.ToString()}:C{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna D
                                _RangeN = (Excel.Range)_xlSht.get_Range($"D{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0";
                                _RangeN.Formula = $"=SUM(D{_IndiceI.ToString()}:D{(_IndiceI + _IndiceF - 1).ToString()})";
                                if (_sAnexo == "ANEXO 7")
                                {
                                    //Columna K
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"K{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(K{_IndiceI.ToString()}:K{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna R
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"R{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(R{_IndiceI.ToString()}:R{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna Y
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"Y{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(Y{_IndiceI.ToString()}:Y{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AE
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AE{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AE{_IndiceI.ToString()}:AE{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AF
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AF{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AF{_IndiceI.ToString()}:AF{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AG
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AG{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AG{_IndiceI.ToString()}:AG{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AH
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AH{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AH{_IndiceI.ToString()}:AH{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AI
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AI{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AI{_IndiceI.ToString()}:AI{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AJ
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AJ{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AJ{_IndiceI.ToString()}:AJ{(_IndiceI + _IndiceF - 1).ToString()})";
                                }
                                _IndiceI = 0;
                                _IndiceF = 0;
                            }

                            _Range = (Excel.Range)xlSht.get_Range($"C{y.ToString()}");
                            _RangeN = (Excel.Range)_xlSht.get_Range($"D{y.ToString()}");
                            string _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                            //ANEXO 7
                            if (_sAnexo == "ANEXO 7")
                            {
                                //Columna K
                                _Range = (Excel.Range)xlSht.get_Range($"J{y.ToString()}");
                                _RangeN = (Excel.Range)_xlSht.get_Range($"K{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna R
                                _Range = (Excel.Range)xlSht.get_Range($"Q{y.ToString()}");
                                _RangeN = (Excel.Range)_xlSht.get_Range($"R{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna Y
                                _Range = (Excel.Range)xlSht.get_Range($"X{y.ToString()}");
                                _RangeN = (Excel.Range)_xlSht.get_Range($"Y{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna AF
                                _RangeN = (Excel.Range)_xlSht.get_Range($"AF{(y).ToString()}");
                                _Range = (Excel.Range)xlSht.get_Range($"AE{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna AH
                                _RangeN = (Excel.Range)_xlSht.get_Range($"AH{(y).ToString()}");
                                _Range = (Excel.Range)xlSht.get_Range($"AG{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                                //Columna AJ
                                _RangeN = (Excel.Range)_xlSht.get_Range($"AJ{(y).ToString()}");
                                _Range = (Excel.Range)xlSht.get_Range($"AI{y.ToString()}");
                                _rValor = _Range.Value;
                                _rValor = _rValor.Replace(",", "");
                                if (_rValor != "")
                                {
                                    if (Int64.TryParse(_rValor, out _Resultado))
                                    {
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                    }
                                    else
                                    {
                                        _RangeN.Value = _rValor;
                                    }
                                }
                            }
                        }
                    }
                    if (_Range.Value == "EXPLICACION")
                    {
                        if (_IndiceF > 0)
                        {
                            if (Int64.TryParse(_RangeS.Value, out _Valor))
                            {
                                _Valor = Convert.ToInt64(_RangeS.Value);
                                _sValor = Convert.ToString(_RangeS.Value);
                                _Hijo = Convert.ToInt32(_sValor.Substring(_sValor.Length - 4, 4));

                                if (_Hijo < 100)
                                {
                                    //Columna C
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"C{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0";
                                    _RangeN.Formula = $"=SUM(C{_IndiceI.ToString()}:C{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna D
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"D{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0";
                                    _RangeN.Formula = $"=SUM(D{_IndiceI.ToString()}:D{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //ANEXO 7
                                    if (_sAnexo == "ANEXO 7")
                                    {
                                        //Columna K
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"K{(y - _IndiceF - 1).ToString()}");
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Formula = $"=SUM(K{_IndiceI.ToString()}:K{(_IndiceI + _IndiceF - 1).ToString()})";
                                        //Columna R
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"R{(y - _IndiceF - 1).ToString()}");
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Formula = $"=SUM(R{_IndiceI.ToString()}:R{(_IndiceI + _IndiceF - 1).ToString()})";
                                        //Columna Y
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"Y{(y - _IndiceF - 1).ToString()}");
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Formula = $"=SUM(Y{_IndiceI.ToString()}:Y{(_IndiceI + _IndiceF - 1).ToString()})";
                                        //Columna AE
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"AE{(_IndiceI - 1).ToString()}");
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Formula = $"=SUM(AE{_IndiceI.ToString()}:AE{(_IndiceI + _IndiceF - 1).ToString()})";
                                        //Columna AF
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"AF{(_IndiceI - 1).ToString()}");
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Formula = $"=SUM(AF{_IndiceI.ToString()}:AF{(_IndiceI + _IndiceF - 1).ToString()})";
                                        //Columna AG
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"AG{(_IndiceI - 1).ToString()}");
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Formula = $"=SUM(AG{_IndiceI.ToString()}:AG{(_IndiceI + _IndiceF - 1).ToString()})";
                                        //Columna AH
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"AH{(_IndiceI - 1).ToString()}");
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Formula = $"=SUM(AH{_IndiceI.ToString()}:AH{(_IndiceI + _IndiceF - 1).ToString()})";
                                        //Columna AI
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"AI{(_IndiceI - 1).ToString()}");
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Formula = $"=SUM(AI{_IndiceI.ToString()}:AI{(_IndiceI + _IndiceF - 1).ToString()})";
                                        //Columna AJ
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"AJ{(_IndiceI - 1).ToString()}");
                                        _RangeN.NumberFormat = "0.00";
                                        _RangeN.Formula = $"=SUM(AJ{_IndiceI.ToString()}:AJ{(_IndiceI + _IndiceF - 1).ToString()})";
                                    }

                                    _IndiceI = 0;
                                    _IndiceF = 0;
                                }
                                else
                                {
                                    _IndiceF += 1;
                                }
                            }
                            else
                            {
                                //Columna C
                                _RangeN = (Excel.Range)_xlSht.get_Range($"C{(y - _IndiceF - 1).ToString()}");
                                _RangeN.NumberFormat = "0";
                                _RangeN.Formula = $"=SUM(C{_IndiceI.ToString()}:C{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna D
                                _RangeN = (Excel.Range)_xlSht.get_Range($"D{(y - _IndiceF - 1).ToString()}");
                                _RangeN.NumberFormat = "0";
                                _RangeN.Formula = $"=SUM(D{_IndiceI.ToString()}:D{(_IndiceI + _IndiceF - 1).ToString()})";
                                //ANEXO 7
                                if (_sAnexo == "ANEXO 7")
                                {
                                    //Columna K
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"K{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(K{_IndiceI.ToString()}:K{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna R
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"R{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(R{_IndiceI.ToString()}:R{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna Y
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"Y{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(Y{_IndiceI.ToString()}:Y{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AE
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AE{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AE{_IndiceI.ToString()}:AE{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AF
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AF{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AF{_IndiceI.ToString()}:AF{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AG
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AG{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AG{_IndiceI.ToString()}:AG{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AH
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AH{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AH{_IndiceI.ToString()}:AH{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AI
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AI{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AI{_IndiceI.ToString()}:AI{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna AJ
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"AJ{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(AJ{_IndiceI.ToString()}:AJ{(_IndiceI + _IndiceF - 1).ToString()})";
                                }

                                _IndiceI = 0;
                                _IndiceF = 0;
                            }

                        }

                        _RangeN = _xlSht.get_Range(string.Format("{0}:{0}", (y), Type.Missing));
                        _RangeN.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        _RangeN = (Excel.Range)_xlSht.get_Range($"A{(y).ToString()}");
                        _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                        _RangeN.Value = _Range.Value;
                        _RangeN = (Excel.Range)_xlSht.get_Range($"B{(y).ToString()}");
                        _Range = (Excel.Range)xlSht.get_Range($"B{y.ToString()}");
                        _RangeN.Value = _Range.Value;
                    }
                }
                if (_Vacio > 11)
                {
                    y = _final + 1;
                }
            }
        }

        private void TransferAnexo2(Excel.Workbook wb, Excel.Worksheet _xAnexo, string _sAnexo, int _inicial, int _final)
        {
            Excel.Range _Range;
            Excel.Range _RangeP;
            Excel.Range _RangeN;
            Excel.Range _RangeS;
            string _sValor;
            Int64 _Valor;
            int _Hijo;
            Int64 _Resultado;
            int _Vacio = 0, _IndiceI = 0, _IndiceF = 0;
            Excel.Worksheet xlSht = _xAnexo;
            Excel.Worksheet _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);

            for (int y = _inicial; y <= _final; y++)
            {
                _RangeP = (Excel.Range)xlSht.get_Range($"A{(y - 1).ToString()}");
                _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                _RangeS = (Excel.Range)xlSht.get_Range($"A{(y + 1).ToString()}");

                if (_Range.Value == "")
                {
                    _Vacio += 1;
                }
                if (_Range.Value != "")
                {
                    _Vacio = 0;
                    if (Int64.TryParse(_Range.Value, out _Valor))
                    {
                        _Valor = Convert.ToInt64(_Range.Value);
                        _sValor = Convert.ToString(_Range.Value);
                        _Hijo = Convert.ToInt32(_sValor.Substring(_sValor.Length - 4, 4));

                        if (_Hijo >= 100)
                        {
                            if (_IndiceF == 0)
                            { _IndiceI = y; }
                            _IndiceF += 1;
                            
                            _RangeN = _xlSht.get_Range(string.Format("{0}:{0}", (y), Type.Missing));
                            _RangeN.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                            _RangeN = (Excel.Range)_xlSht.get_Range($"A{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                            _RangeN.Value = _Range.Value;
                            Generales.AddNamedRange(y, 1, "IA_0" + _RangeN.Value,_xlSht);
                            _RangeN = (Excel.Range)_xlSht.get_Range($"B{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"B{y.ToString()}");
                            _RangeN.Value = _Range.Value;
                            //Columna F
                            _RangeN = (Excel.Range)_xlSht.get_Range($"F{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"C{y.ToString()}");
                            string _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                            //Columna G
                            _RangeN = (Excel.Range)_xlSht.get_Range($"G{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"D{y.ToString()}");
                            _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                            //Columna H
                            _RangeN = (Excel.Range)_xlSht.get_Range($"H{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"E{y.ToString()}");
                            _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                        }
                        else
                        {
                            if (_IndiceF > 0)
                            {
                                //Columna C
                                _RangeN = (Excel.Range)_xlSht.get_Range($"C{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0";
                                _RangeN.Formula = $"=SUM(C{_IndiceI.ToString()}:C{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna D
                                _RangeN = (Excel.Range)_xlSht.get_Range($"D{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0";
                                _RangeN.Formula = $"=SUM(D{_IndiceI.ToString()}:D{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna E
                                _RangeN = (Excel.Range)_xlSht.get_Range($"E{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(E{_IndiceI.ToString()}:E{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna F
                                _RangeN = (Excel.Range)_xlSht.get_Range($"F{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(F{_IndiceI.ToString()}:F{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna G
                                _RangeN = (Excel.Range)_xlSht.get_Range($"G{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(G{_IndiceI.ToString()}:G{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna H
                                _RangeN = (Excel.Range)_xlSht.get_Range($"H{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(H{_IndiceI.ToString()}:H{(_IndiceI + _IndiceF - 1).ToString()})";

                                _IndiceI = 0;
                                _IndiceF = 0;
                            }

                            //Columna F
                            _Range = (Excel.Range)xlSht.get_Range($"C{y.ToString()}");
                            _RangeN = (Excel.Range)_xlSht.get_Range($"F{y.ToString()}");
                            string _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                            //Columna G
                            _Range = (Excel.Range)xlSht.get_Range($"D{y.ToString()}");
                            _RangeN = (Excel.Range)_xlSht.get_Range($"G{y.ToString()}");
                            _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                            //Columna H
                            _Range = (Excel.Range)xlSht.get_Range($"E{y.ToString()}");
                            _RangeN = (Excel.Range)_xlSht.get_Range($"H{y.ToString()}");
                            _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                        }
                    }
                    if (_Range.Value == "EXPLICACION")
                    {
                        if (_IndiceF > 0)
                        {
                            if (Int64.TryParse(_RangeS.Value, out _Valor))
                            {
                                _Valor = Convert.ToInt64(_RangeS.Value);
                                _sValor = Convert.ToString(_RangeS.Value);
                                _Hijo = Convert.ToInt32(_sValor.Substring(_sValor.Length - 4, 4));

                                if (_Hijo < 100)
                                {
                                    //Columna C
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"C{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0";
                                    _RangeN.Formula = $"=SUM(C{_IndiceI.ToString()}:C{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna D
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"D{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0";
                                    _RangeN.Formula = $"=SUM(D{_IndiceI.ToString()}:D{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna E
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"E{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(E{_IndiceI.ToString()}:E{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna F
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"F{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(F{_IndiceI.ToString()}:F{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna G
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"G{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(G{_IndiceI.ToString()}:G{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna H
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"H{(_IndiceI - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(H{_IndiceI.ToString()}:H{(_IndiceI + _IndiceF - 1).ToString()})";

                                    _IndiceI = 0;
                                    _IndiceF = 0;
                                }
                                else
                                {
                                    _IndiceF += 1;
                                }
                            }
                        }

                        _RangeN = _xlSht.get_Range(string.Format("{0}:{0}", (y), Type.Missing));
                        _RangeN.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        _RangeN = (Excel.Range)_xlSht.get_Range($"A{(y).ToString()}");
                        _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                        _RangeN.Value = _Range.Value;
                        _RangeN = (Excel.Range)_xlSht.get_Range($"B{(y).ToString()}");
                        _Range = (Excel.Range)xlSht.get_Range($"B{y.ToString()}");
                        _RangeN.Value = _Range.Value;
                    }
                }
                if (_Vacio > 11)
                {
                    y = _final + 1;
                }
            }
        }

        private void TransferAnexo3(Excel.Workbook wb, Excel.Worksheet _xAnexo, string _sAnexo, int _inicial, int _final)
        {
            Excel.Range _Range;
            Excel.Range _RangeB;
            Excel.Range _RangeP;
            Excel.Range _RangeN;
            Excel.Range _RangeS;
            string _sValor;
            Int64 _Valor;
            int _Hijo;
            Int64 _Resultado;
            int _Vacio = 0, _IndiceI = 0, _IndiceF = 0;
            Excel.Worksheet xlSht = _xAnexo;
            Excel.Worksheet _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);

            for (int y = _inicial; y <= _final; y++)
            {
                _RangeP = (Excel.Range)xlSht.get_Range($"A{(y - 1).ToString()}");
                _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                _RangeB = (Excel.Range)xlSht.get_Range($"B{y.ToString()}");
                _RangeS = (Excel.Range)xlSht.get_Range($"A{(y + 1).ToString()}");

                if (_Range.Value == "")
                {
                    _Vacio += 1;
                }
                if (_Range.Value != "")
                {
                    string _sRangeB = Convert.ToString(_RangeB.Value);
                    int _RenglonB = 0;

                    if (_sRangeB != "")
                    {
                        _sRangeB = _sRangeB == null ? "0" : _sRangeB;
                        _RenglonB = Convert.ToString(_sRangeB).IndexOf("2015");
                    }
                    _Vacio = 0;

                    if (_RenglonB > 0)
                    {
                        if (Int64.TryParse(_Range.Value, out _Valor))
                        {
                            _Valor = Convert.ToInt64(_Range.Value);
                            _sValor = Convert.ToString(_Range.Value);
                            _Hijo = Convert.ToInt32(_sValor.Substring(_sValor.Length - 4, 4));

                            if (_Hijo >= 100)
                            {
                                if (_IndiceF == 0)
                                { _IndiceI = y; }
                                _IndiceF += 1;

                                for (int z = _inicial; z <= _final; z++)
                                {
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"A{z.ToString()}");

                                    if (Convert.ToString(_RangeN.Value) == Convert.ToString(_RangeP.Value))
                                    {
                                        //Columna Template Nuevo
                                        _RangeN = _xlSht.get_Range(string.Format("{0}:{0}", (z + 1), Type.Missing));
                                        _RangeN.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                                        Generales.AddNamedRange(z + 1, 1, "IA_0" + _RangeN.Value, _xlSht);

                                        for (int a = 3; a <= 21; a++)
                                        {
                                            //Columna XSPR
                                            _Range = (Excel.Range)xlSht.get_Range($"{Generales.ColumnAdress(a)}{y.ToString()}");
                                            //Columna Template Nuevo
                                            _RangeN = (Excel.Range)_xlSht.get_Range($"{Generales.ColumnAdress(a)}{(z + 1).ToString()}");
                                            //Valores
                                            string _rValor = _Range.Value;
                                            _rValor = _rValor.Replace(",", "");
                                            if (_rValor != "")
                                            {
                                                if (Int64.TryParse(_rValor, out _Resultado))
                                                {
                                                    if (a > 2)
                                                    {
                                                        _RangeN.NumberFormat = "0";
                                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                                    }
                                                }
                                                else
                                                {
                                                    _RangeN.Value = _rValor;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (_IndiceF > 0)
                                {
                                    //Columna XSPR
                                    _Range = (Excel.Range)xlSht.get_Range($"A{(_IndiceI - 1).ToString()}");

                                    for (int z = _inicial; z <= _final; z++)
                                    {
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"A{z.ToString()}");

                                        if (Convert.ToString(_RangeN.Value) == Convert.ToString(_Range.Value))
                                        {
                                            for (int a = 3; a <= 21; a++)
                                            {
                                                //Columnas
                                                _RangeN = (Excel.Range)_xlSht.get_Range($"{Generales.ColumnAdress(a)}{(_IndiceI - 1).ToString()}");
                                                _RangeN.NumberFormat = "0";
                                                _RangeN.Formula = $"=SUM({Generales.ColumnAdress(a)}{_IndiceI.ToString()}:{Generales.ColumnAdress(a)}{(_IndiceI + _IndiceF - 1).ToString()})";
                                            }
                                        }
                                    }

                                    _IndiceI = 0;
                                    _IndiceF = 0;
                                }

                                for (int z = _inicial; z <= _final; z++)
                                {
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"A{z.ToString()}");

                                    if (Convert.ToString(_RangeN.Value) == Convert.ToString(_Range.Value))
                                    {
                                        for (int a = 3; a <= 21; a++)
                                        {
                                            //Columna XSPR
                                            _Range = (Excel.Range)xlSht.get_Range($"{Generales.ColumnAdress(a)}{y.ToString()}");
                                            //Columna Template Nuevo
                                            _RangeN = (Excel.Range)_xlSht.get_Range($"{Generales.ColumnAdress(a)}{z.ToString()}");
                                            //Valores
                                            string _rValor = _Range.Value;
                                            _rValor = _rValor.Replace(",", "");
                                            if (_rValor != "")
                                            {
                                                if (Int64.TryParse(_rValor, out _Resultado))
                                                {
                                                    if (a > 2)
                                                    {
                                                        _RangeN.NumberFormat = "0";
                                                        _RangeN.Value = Convert.ToInt64(_rValor);
                                                    }
                                                }
                                                else
                                                {
                                                    _RangeN.Value = _rValor;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (_Range.Value == "EXPLICACION")
                        {
                            if (_IndiceF > 0)
                            {
                                if (Int64.TryParse(_RangeS.Value, out _Valor))
                                {
                                    _Valor = Convert.ToInt64(_RangeS.Value);
                                    _sValor = Convert.ToString(_RangeS.Value);
                                    _Hijo = Convert.ToInt32(_sValor.Substring(_sValor.Length - 4, 4));

                                    if (_Hijo < 100)
                                    {
                                        //Columna XSPR
                                        _Range = (Excel.Range)xlSht.get_Range($"A{(_IndiceI - 1).ToString()}");

                                        for (int z = _inicial; z <= _final; z++)
                                        {
                                            _RangeN = (Excel.Range)_xlSht.get_Range($"A{z.ToString()}");

                                            if (Convert.ToString(_RangeN.Value) == Convert.ToString(_Range.Value))
                                            {
                                                for (int a = 3; a <= 21; a++)
                                                {
                                                    //Columnas
                                                    _RangeN = (Excel.Range)_xlSht.get_Range($"{Generales.ColumnAdress(a)}{(_IndiceI - 1).ToString()}");
                                                    _RangeN.NumberFormat = "0";
                                                    _RangeN.Formula = $"=SUM({Generales.ColumnAdress(a)}{_IndiceI.ToString()}:{Generales.ColumnAdress(a)}{(_IndiceI + _IndiceF - 1).ToString()})";
                                                }
                                            }
                                        }

                                        _IndiceI = 0;
                                        _IndiceF = 0;
                                    }
                                    else
                                    {
                                        _IndiceF += 1;
                                    }
                                }
                            }

                            for (int z = _inicial; z <= _final; z++)
                            {
                                _RangeN = (Excel.Range)_xlSht.get_Range($"A{z.ToString()}");

                                if (Convert.ToString(_RangeN.Value) == Convert.ToString(_RangeP.Value))
                                {
                                    //Columna Template Nuevo
                                    _RangeN = _xlSht.get_Range(string.Format("{0}:{0}", (z + 1), Type.Missing));
                                    _RangeN.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                                    for (int a = 1; a <= 21; a++)
                                    {
                                        //Columna XSPR
                                        _Range = (Excel.Range)xlSht.get_Range($"{Generales.ColumnAdress(a)}{y.ToString()}");
                                        _RangeN = (Excel.Range)_xlSht.get_Range($"{Generales.ColumnAdress(a)}{(z + 1).ToString()}");
                                        //Valores
                                        _RangeN.Value = _Range.Value;
                                    }
                                }
                            }
                        }
                    }
                }
                if (_Vacio > 11)
                {
                    y = _final + 1;
                }
            }
        }

        private void TransferAnexo8(Excel.Workbook wb, Excel.Worksheet _xAnexo, string _sAnexo, int _inicial, int _final)
        {
            Excel.Range _Range;
            Excel.Range _RangeP;
            Excel.Range _RangeN;
            Excel.Range _RangeS;
            string _sValor;
            Int64 _Valor;
            int _Hijo;
            Int64 _Resultado;
            int _Vacio = 0, _IndiceI = 0, _IndiceF = 0;
            Excel.Worksheet xlSht = _xAnexo;
            Excel.Worksheet _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);

            for (int y = _inicial; y <= _final; y++)
            {
                _RangeP = (Excel.Range)xlSht.get_Range($"A{(y - 1).ToString()}");
                _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                _RangeS = (Excel.Range)xlSht.get_Range($"A{(y + 1).ToString()}");

                if (_Range.Value == "")
                {
                    _Vacio += 1;
                }
                if (_Range.Value != "")
                {
                    _Vacio = 0;
                    if (Int64.TryParse(_Range.Value, out _Valor))
                    {
                        _Valor = Convert.ToInt64(_Range.Value);
                        _sValor = Convert.ToString(_Range.Value);
                        _Hijo = Convert.ToInt32(_sValor.Substring(_sValor.Length - 4, 4));

                        if (_Hijo >= 100)
                        {
                            if (_IndiceF == 0)
                            { _IndiceI = y; }
                            _IndiceF += 1;

                            _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);
                            _RangeN = _xlSht.get_Range(string.Format("{0}:{0}", (y), Type.Missing));
                            _RangeN.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                            _RangeN = (Excel.Range)_xlSht.get_Range($"A{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                            _RangeN.Value = _Range.Value;
                            Generales.AddNamedRange(y, 1, "IA_0" + _RangeN.Value, _xlSht);
                            _RangeN = (Excel.Range)_xlSht.get_Range($"B{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"B{y.ToString()}");
                            _RangeN.Value = _Range.Value;
                            //Columna D
                            _RangeN = (Excel.Range)_xlSht.get_Range($"D{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"C{y.ToString()}");
                            string _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                            //Columna F
                            _RangeN = (Excel.Range)_xlSht.get_Range($"F{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"E{y.ToString()}");
                            _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                            //Columna H
                            _RangeN = (Excel.Range)_xlSht.get_Range($"H{(y).ToString()}");
                            _Range = (Excel.Range)xlSht.get_Range($"G{y.ToString()}");
                            _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                        }
                        else
                        {
                            if (_IndiceF > 0)
                            {
                                _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);
                                //Columna C
                                _RangeN = (Excel.Range)_xlSht.get_Range($"C{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(C{_IndiceI.ToString()}:C{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna D
                                _RangeN = (Excel.Range)_xlSht.get_Range($"D{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(D{_IndiceI.ToString()}:D{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna E
                                _RangeN = (Excel.Range)_xlSht.get_Range($"E{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(E{_IndiceI.ToString()}:E{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna F
                                _RangeN = (Excel.Range)_xlSht.get_Range($"F{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(F{_IndiceI.ToString()}:F{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna G
                                _RangeN = (Excel.Range)_xlSht.get_Range($"G{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(G{_IndiceI.ToString()}:G{(_IndiceI + _IndiceF - 1).ToString()})";
                                //Columna H
                                _RangeN = (Excel.Range)_xlSht.get_Range($"H{(_IndiceI - 1).ToString()}");
                                _RangeN.NumberFormat = "0.00";
                                _RangeN.Formula = $"=SUM(H{_IndiceI.ToString()}:H{(_IndiceI + _IndiceF - 1).ToString()})";

                                _IndiceI = 0;
                                _IndiceF = 0;
                            }

                            //Columna D
                            _Range = (Excel.Range)xlSht.get_Range($"C{y.ToString()}");
                            _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);
                            _RangeN = (Excel.Range)_xlSht.get_Range($"D{y.ToString()}");
                            string _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                            //Columna F
                            _Range = (Excel.Range)xlSht.get_Range($"E{y.ToString()}");
                            _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);
                            _RangeN = (Excel.Range)_xlSht.get_Range($"F{y.ToString()}");
                            _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                            //Columna H
                            _Range = (Excel.Range)xlSht.get_Range($"G{y.ToString()}");
                            _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);
                            _RangeN = (Excel.Range)_xlSht.get_Range($"H{y.ToString()}");
                            _rValor = _Range.Value;
                            _rValor = _rValor.Replace(",", "");
                            if (_rValor != "")
                            {
                                if (Int64.TryParse(_rValor, out _Resultado))
                                {
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Value = Convert.ToInt64(_rValor);
                                }
                                else
                                {
                                    _RangeN.Value = _rValor;
                                }
                            }
                        }
                    }
                    if (_Range.Value == "EXPLICACION")
                    {
                        if (_IndiceF > 0)
                        {
                            if (Int64.TryParse(_RangeS.Value, out _Valor))
                            {
                                _Valor = Convert.ToInt64(_RangeS.Value);
                                _sValor = Convert.ToString(_RangeS.Value);
                                _Hijo = Convert.ToInt32(_sValor.Substring(_sValor.Length - 4, 4));

                                if (_Hijo < 100)
                                {
                                    _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);
                                    //Columna C
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"C{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(C{_IndiceI.ToString()}:C{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna D
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"D{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(D{_IndiceI.ToString()}:D{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna E
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"E{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(E{_IndiceI.ToString()}:E{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna F
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"F{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(F{_IndiceI.ToString()}:F{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna G
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"G{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(G{_IndiceI.ToString()}:G{(_IndiceI + _IndiceF - 1).ToString()})";
                                    //Columna H
                                    _RangeN = (Excel.Range)_xlSht.get_Range($"H{(y - _IndiceF - 1).ToString()}");
                                    _RangeN.NumberFormat = "0.00";
                                    _RangeN.Formula = $"=SUM(H{_IndiceI.ToString()}:H{(_IndiceI + _IndiceF - 1).ToString()})";

                                    _IndiceI = 0;
                                    _IndiceF = 0;
                                }
                                else
                                {
                                    _IndiceF += 1;
                                }
                            }
                        }

                        _xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_sAnexo);
                        _RangeN = _xlSht.get_Range(string.Format("{0}:{0}", (y), Type.Missing));
                        _RangeN.Insert(Excel.XlInsertShiftDirection.xlShiftDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                        _RangeN = (Excel.Range)_xlSht.get_Range($"A{(y).ToString()}");
                        _Range = (Excel.Range)xlSht.get_Range($"A{y.ToString()}");
                        _RangeN.Value = _Range.Value;
                        _RangeN = (Excel.Range)_xlSht.get_Range($"B{(y).ToString()}");
                        _Range = (Excel.Range)xlSht.get_Range($"B{y.ToString()}");
                        _RangeN.Value = _Range.Value;
                    }
                }
                if (_Vacio > 11)
                {
                    y = _final + 1;
                }
            }
        }

        private void fnProgressBar(int _Progress)
        {
            this.pgbFile.Minimum = 0;
            this.pgbFile.Maximum = 100;
            Invoke(new System.Action(() => this.pgbFile.Value = _Progress));
        }

        private void ConversionMasiva_Load(object sender, EventArgs e)
        {
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            string _Path = Configuration.Path;
            string _Anexo = "";
            int _IdTipo = 1;
            oComprobacion[] _Comprobaciones;
            
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
                        this.btnConvertir.Enabled = false;
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

            //Libro Actual de Excel.
            Excel.Worksheet xlSht;
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            FileInfo _Excel = new FileInfo(wb.FullName);

            if (_Excel.Extension != ".xspr")
            {
                this.cmbAnio.Enabled = false;
                this.chkbAll.Enabled = false;
                this.btnCarpeta.Enabled = false;
                this.btnConvertir.Enabled = false;
                return;
            }
            else
            {
                //Combo de los Años.
                FillYears(cmbAnio);
                //Check para seleccionar todos los Anexos.
                chkbAll.Checked = true;
                // el nombre de una Key debe incluir un root valido.
                const string userRoot = "HKEY_CURRENT_USER";
                const string subkey = "Software\\Microsoft\\Office\\Excel\\Addins\\SAT.Dictamenes.SIPRED.Client";
                const string keyName = userRoot + "\\" + subkey;

                object addInName = "SAT.Dictamenes.SIPRED.Client";
                Registry.SetValue(keyName, "LoadBehavior", 0);
                Globals.ThisAddIn.Application.COMAddIns.Item(ref addInName).Connect = false;
                Globals.ThisAddIn.Application.Visible = true;
                Excel.Range _Range;

                try
                {
                    xlSht = (Excel.Worksheet)wb.Worksheets.get_Item("Contribuyente");
                    _Range = (Excel.Range)xlSht.get_Range("C9");
                    this.txtRFC.Text = Convert.ToString(_Range.Value);
                    this.txtTipo.Text = "ESTADOS FINANCIEROS GENERAL";
                }
                catch (Exception ex)
                {
                    _Error = true;
                    _sError = $"El archivo XSPR fue cerrado, al activar el módulo de convertir y mostrar el mensaje de guardar el XSPR aseguresé de darle cilc en [Cancelar] para tener activo el XSPR.: {ex.Message.ToString()}";
                    return;
                }

                _Comprobaciones = Assembler.LoadJson<oComprobacion[]>($"{_Path}\\jsons\\Comprobaciones.json");

                foreach (oComprobacion _Comprobacion in _Comprobaciones.Where(o => o.IdTipoPlantilla == Convert.ToInt32(_IdTipo)).ToArray())
                {
                    xlSht = (Excel.Worksheet)wb.Worksheets.get_Item(_Comprobacion.Destino.Anexo);
                    _Range = (Excel.Range)xlSht.get_Range("B1");

                    if (_Anexo != _Comprobacion.Destino.Anexo)
                    {
                        switch (_Comprobacion.Destino.Anexo)
                        {
                            case "ANEXO 1":
                                fnGridAnexo(_Comprobacion.Destino.Anexo, _Range.Value, true);
                                break;
                            case "ANEXO 2":
                                fnGridAnexo(_Comprobacion.Destino.Anexo, _Range.Value, true);
                                break;
                            case "ANEXO 3":
                                fnGridAnexo(_Comprobacion.Destino.Anexo, _Range.Value, true);
                                break;
                            case "ANEXO 4":
                                fnGridAnexo(_Comprobacion.Destino.Anexo, _Range.Value, true);
                                break;
                            case "ANEXO 5":
                                fnGridAnexo(_Comprobacion.Destino.Anexo, _Range.Value, true);
                                break;
                            case "ANEXO 6":
                                fnGridAnexo(_Comprobacion.Destino.Anexo, _Range.Value, true);
                                break;
                            case "ANEXO 7":
                                fnGridAnexo(_Comprobacion.Destino.Anexo, _Range.Value, true);
                                break;
                            case "ANEXO 8":
                                fnGridAnexo(_Comprobacion.Destino.Anexo, _Range.Value, true);
                                break;
                        }
                        _Anexo = _Comprobacion.Destino.Anexo;
                    }
                }

                dgvAnexos.DataSource = dtAnexo;
                dgvAnexos.Columns[0].Visible = false;
                dgvAnexos.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
                dgvAnexos.Columns[1].Width = 240;
                dgvAnexos.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                dgvAnexos.Columns[2].Width = 100;
                dgvAnexos.Columns[2].HeaderText = "";
                this.btnConvertir.Focus();
            }
        }

        private void fnGridAnexo(string _Anexo, string _Descripcion, bool _Convertir)
        {
            DataRow drAnexo;
            drAnexo = dtAnexo.NewRow();
            drAnexo["Anexo"] = _Anexo;
            drAnexo["Descripcion"] = _Descripcion;
            drAnexo["Convertir"] = _Convertir;
            dtAnexo.Rows.Add(drAnexo);
        }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCarpeta_Click(object sender, EventArgs e)
        {
            string _DestinationPath = "";

            fbdCarpeta.ShowDialog();
            _DestinationPath = fbdCarpeta.SelectedPath;

            if (_DestinationPath == "")
            {
                MessageBox.Show("Debe especificar un ruta", "Ruta Invalida", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            this.txtRuta.Text = _DestinationPath;
        }

        private void dgvAnexos_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(e.ColumnIndex ==2)
            {
                if (Convert.ToBoolean(dtAnexo.Rows[e.RowIndex]["Convertir"]) == true)
                dtAnexo.Rows[e.RowIndex]["Convertir"] = false;
                else
                    dtAnexo.Rows[e.RowIndex]["Convertir"] = true;
            }

            dgvAnexos.DataSource = dtAnexo;
            dgvAnexos.Columns[0].Visible = false;
            dgvAnexos.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
            dgvAnexos.Columns[1].Width = 240;
            dgvAnexos.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            dgvAnexos.Columns[2].Width = 100;
            dgvAnexos.Columns[2].HeaderText = "";
        }

        private void chkbAll_CheckedChanged(object sender, EventArgs e)
        {
            int _Rows = dtAnexo.Rows.Count;

            for(int x=0; x<_Rows; x++)
            {
                dtAnexo.Rows[x]["Convertir"] = false;
            }

            if(chkbAll.Checked == true)
            {
                for (int x = 0; x < _Rows; x++)
                {
                    dtAnexo.Rows[x]["Convertir"] = true;
                }
            }
            dgvAnexos.DataSource = dtAnexo;
            dgvAnexos.Columns[0].Visible = false;
            dgvAnexos.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft;
            dgvAnexos.Columns[1].Width = 240;
            dgvAnexos.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
            dgvAnexos.Columns[2].Width = 100;
            dgvAnexos.Columns[2].HeaderText = "";
        }

        private void ConversionMasiva_Shown(object sender, EventArgs e)
        {
            //Libro Actual de Excel.
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            FileInfo _Excel = new FileInfo(wb.FullName);

            this.cmbAnio.Enabled = true;
            this.chkbAll.Enabled = true;
            this.btnCarpeta.Enabled = true;
            this.btnConvertir.Enabled = true;

            if (_Excel.Extension != ".xspr")
            {
                MessageBox.Show("Archivo no válido, favor de generar el archivo mediante el AddIn de SIPRED", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            if(_Error)
            {
                MessageBox.Show(_sError, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
        }
    }
}
