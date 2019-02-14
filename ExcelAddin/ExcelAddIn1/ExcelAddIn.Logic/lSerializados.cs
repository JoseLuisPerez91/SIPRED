using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;
using ExcelAddIn.Objects;
using ExcelAddIn.Access;
using System.Net;

namespace ExcelAddIn.Logic {
    public class lSerializados : aSerializados {
        public lSerializados() { }

        public KeyValuePair<bool, string[]> ObtenerSerializados() {
            bool _Key= true;
            _Messages = new List<string>();

            if (!CheckConnection("http://www.google.com"))
            {
                string[] input = { "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión." };
                _Messages.AddRange(input);
                _Key = false;
            }
            else
            {
                KeyValuePair<bool, string[]> _TiposPlantillas = ObtenerTiposPlantillas();
                KeyValuePair<bool, string[]> _Cruces = ObtenerCruces();
                KeyValuePair<bool, string[]> _Plantillas = ObtenerPlantillas();
                KeyValuePair<bool, string[]> _Comprobaciones = ObtenerComprobaciones();
                KeyValuePair<bool, string[]> _Validaciones = ObtenerValidacionCruces();
                KeyValuePair<bool, string[]> _Indices = ObtenerIndices();
                _Key = (!_TiposPlantillas.Key || !_Cruces.Key || !_Plantillas.Key || !_Comprobaciones.Key || !_Validaciones.Key || !_Indices.Key);
                _Messages.AddRange(_TiposPlantillas.Value);
                _Messages.AddRange(_Cruces.Value);
                _Messages.AddRange(_Plantillas.Value);
                _Messages.AddRange(_Comprobaciones.Value);
                _Messages.AddRange(_Validaciones.Value);
                _Messages.AddRange(_Indices.Value);
            }
            return new KeyValuePair<bool, string[]>(_Key, _Messages.ToArray());
        }
        new KeyValuePair<bool, string[]> ObtenerIndices()
        {
            KeyValuePair<KeyValuePair<bool, string>, object> _result = base.ObtenerIndices();
            if (_result.Key.Key)
            {
                string _JsonData = (string)_result.Value;
                File.WriteAllText($"{Access.Configuration.Path}\\jsons\\Indices.json", _JsonData);
                return new KeyValuePair<bool, string[]>(true, new string[] { "Se generó correctamente el archivo json para los Índices." });
            }
            else { }
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }

        new KeyValuePair<bool, string[]> ObtenerValidacionCruces()
        {
            KeyValuePair<KeyValuePair<bool, string>, object> _result = base.ObtenerValidacionCruces();
            if (_result.Key.Key)
            {
                string _JsonData = (string)_result.Value;
                File.WriteAllText($"{Access.Configuration.Path}\\jsons\\ValidacionCruces.json", _JsonData);
                return new KeyValuePair<bool, string[]>(true, new string[] { "Se generó correctamente el archivo json para la Validacion de Cruces." });
            }
            else { }
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }
        new KeyValuePair<bool, string[]> ObtenerTiposPlantillas() {
            KeyValuePair<KeyValuePair<bool, string>, object> _result = base.ObtenerTiposPlantillas();
            if(_result.Key.Key) {
                string _JsonData = (string)_result.Value;
                File.WriteAllText($"{Access.Configuration.Path}\\jsons\\TiposPlantillas.json", _JsonData);
                return new KeyValuePair<bool, string[]>(true, new string[] { "Se generó correctamente el archivo json para los tipos de plantillas." });
            } else { }
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }

        new KeyValuePair<bool, string[]> ObtenerCruces() {
            KeyValuePair<KeyValuePair<bool, string>, object> _result = base.ObtenerCruces();
            if(_result.Key.Key) {
                string _JsonData = (string)_result.Value;
                if(string.IsNullOrEmpty(_JsonData) || string.IsNullOrWhiteSpace(_JsonData))
                    return new KeyValuePair<bool, string[]>(false, new string[] { "No se encontro información para la generación de los cruces." });
                File.WriteAllText($"{Access.Configuration.Path}\\jsons\\Cruces.json", _JsonData.Replace("\\\"", "\""));
                return new KeyValuePair<bool, string[]>(true, new string[] { "Se generó correctamente el archivo json para los cruces." });
            }
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }

        new KeyValuePair<bool, string[]> ObtenerPlantillas() {
            _Messages = new List<string>();
            KeyValuePair<KeyValuePair<bool, string>, object> _result = base.ObtenerPlantillas();
            if(_result.Key.Key) {
                string _JsonData = (string)_result.Value, _FullPath = $"{Access.Configuration.Path}\\jsons\\Plantillas.json";
                oPlantilla[] _Templates = JsonConvert.DeserializeObject<oPlantilla[]>(_JsonData);
                foreach(oPlantilla _Template in _Templates) {
                    KeyValuePair<KeyValuePair<bool, string>, object> _resultFile = base.ObtenerArchivoPlantilla(_Template.IdPlantilla);
                    if(_resultFile.Key.Key) {
                        byte[] _TemplateFile = (byte[])_resultFile.Value;
                        try {
                            File.WriteAllBytes($"{Access.Configuration.Path}\\templates\\{_Template.Nombre}", _TemplateFile);
                        } catch(Exception _ex) {
                            _Messages.Add(_ex.InnerException?.Message ?? _ex.Message);
                            _Messages.Add(_ex.InnerException?.StackTrace ?? _ex.StackTrace);
                        }
                    }
                }
                if(_Messages.Count() == 0) {
                    File.WriteAllText(_FullPath, _JsonData);
                    _Messages.Add("Se generó correctamente el archivo json para las plantillas.");
                    return new KeyValuePair<bool, string[]>(true, _Messages.ToArray());
                } else {
                    _Messages.Add("Ocurrio un error al momento de generar el archivo json de las Plantillas.");
                    return new KeyValuePair<bool, string[]>(false, _Messages.ToArray());
                }
            }
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }

        new KeyValuePair<bool, string[]> ObtenerComprobaciones() {
            KeyValuePair<KeyValuePair<bool, string>, object> _result = base.ObtenerComprobaciones();
            if(_result.Key.Key) {
                string _JsonData = (string)_result.Value, _FullPath = $"{Access.Configuration.Path}\\jsons\\Comprobaciones.json";
                oComprobacion[] _Comprobaciones = JsonConvert.DeserializeObject<oComprobacion[]>(_JsonData);
                string _JsonComprobaciones = InicializarComprobaciones(_Comprobaciones);
                File.WriteAllText(_FullPath, _JsonComprobaciones);
                return new KeyValuePair<bool, string[]>(true, new string[] { "Se generó correctamente el archivo json de las comprobaciones." });
            }
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }

        string InicializarComprobaciones(oComprobacion[] _Comprobaciones) {
            oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>($"{Access.Configuration.Path}\\jsons\\Plantillas.json");
            foreach(oPlantilla _Template in _Templates) {
                FileInfo _Excel = new FileInfo($"{Access.Configuration.Path}\\templates\\{_Template.Nombre}");
                using(ExcelPackage _package = new ExcelPackage(_Excel)) {
                    foreach(oComprobacion _Comprobacion in _Comprobaciones.Where(o => o.IdTipoPlantilla == _Template.IdTipoPlantilla).ToArray()) {
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
                                _Comprobacion.Destino.setCeldaExcel(_workSheet.Cells[_Comprobacion.Destino.Fila, _Comprobacion.Destino.Columna], "");
                                foreach(oCelda _Celda in _Celdas) {
                                    _Celda.Fila = _Comprobacion.Destino.Fila;
                                    _Celda.setCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna], _Comprobacion.Destino.Anexo);
                                }
                                foreach(oCelda _Celda in _cCeldas) {
                                    _Celda.Fila = _Comprobacion.Destino.Fila;
                                    _Celda.setCeldaExcel(_workSheet.Cells[_Celda.Fila, _Celda.Columna], _Comprobacion.Destino.Anexo);
                                }
                                oCelda[] _Faltantes = _Comprobacion.Celdas.Where(o => o.Fila == -1).ToArray();
                                foreach(oCelda _Faltante in _Faltantes) {
                                    oCelda _Result = _Comprobaciones.Where(o => o.Destino != null && o.Destino.Indice == _Faltante.Indice && o.Destino.Anexo == _Faltante.Anexo.ToUpper()).Select(o => o.Destino).FirstOrDefault();
                                    if(_Result != null) {
                                        _Faltante.Fila = _Result.Fila;
                                        _Faltante.setCeldaExcel(_workSheet.Cells[_Faltante.Fila, _Faltante.Columna], _Comprobacion.Destino.Anexo);
                                    }
                                    if(_Result == null) {
                                        ExcelWorksheet _ws = _package.Workbook.Worksheets[_Faltante.Anexo];
                                        int _mv = _ws.Dimension.Rows + 1;
                                        int _mr = (_ws.Dimension.Rows / 2) + (_ws.Dimension.Rows % 2);
                                        for(int j = 1; j <= _mr; j++) {
                                            _Faltante.Fila = (_ws.Cells[j, 1].Text == _Faltante.Indice) ? j : _Faltante.Fila;
                                            _Faltante.Fila = (_ws.Cells[(_mv - j), 1].Text == _Faltante.Indice) ? _mv - j : _Faltante.Fila;
                                            if(_Faltante.Fila > -1) {
                                                _Faltante.setCeldaExcel(_ws.Cells[_Faltante.Fila, _Faltante.Columna], _Comprobacion.Destino.Anexo);
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
                                        _Faltante.setCeldaExcel(_workSheet.Cells[_Faltante.Fila, _Faltante.Columna], _Comprobacion.Destino.Anexo);
                                    }
                                    if(_Result == null) {
                                        ExcelWorksheet _ws = _package.Workbook.Worksheets[_Faltante.Anexo];
                                        int _mv = _ws.Dimension.Rows + 1;
                                        int _mr = (_ws.Dimension.Rows / 2) + (_ws.Dimension.Rows % 2);
                                        for(int j = 1; j <= _mr; j++) {
                                            _Faltante.Fila = (_ws.Cells[j, 1].Text == _Faltante.Indice) ? j : _Faltante.Fila;
                                            _Faltante.Fila = (_ws.Cells[(_mv - j), 1].Text == _Faltante.Indice) ? _mv - j : _Faltante.Fila;
                                            if(_Faltante.Fila > -1) {
                                                _Faltante.setCeldaExcel(_ws.Cells[_Faltante.Fila, _Faltante.Columna], _Comprobacion.Destino.Anexo);
                                                break;
                                            }
                                        }
                                    }
                                }
                                break;
                            }
                        }
                        _Comprobacion.setFormulaExcel();
                    }
                }
            }
            return JsonConvert.SerializeObject(_Comprobaciones);
        }
        public bool CheckConnection(String URL)
        {
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URL);
                request.Timeout = 5000;
                request.Credentials = CredentialCache.DefaultNetworkCredentials;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                if (response.StatusCode == HttpStatusCode.OK)
                    return true;
                else
                    return false;
            }
            catch
            {
                return false;
            }
        }
    }
}