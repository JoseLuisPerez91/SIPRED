using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.IO;
using ExcelAddIn.Objects;
using ExcelAddIn.Access;

namespace ExcelAddIn.Logic {
    public class lSerializados : aSerializados {
        string _Path = Access.Configuration.Path;

        public lSerializados() { }

        public KeyValuePair<bool, string> ObtenerSerializados() {
            return new KeyValuePair<bool, string>();
        }

        new KeyValuePair<bool, string> ObtenerTiposPlantillas() {
            KeyValuePair<KeyValuePair<bool, string>, object> _result = base.ObtenerTiposPlantillas();
            if(_result.Key.Key) {
                string _JsonData = (string)_result.Value;
                File.WriteAllText($"{_Path}\\jsons\\TiposPlantillas.json", _JsonData);
                return new KeyValuePair<bool, string>(true, "Se generó correctamente el archivo json para los tipos de plantillas.");
            }
            return _result.Key;
        }

        new KeyValuePair<bool, string> ObtenerCruces() {
            KeyValuePair<KeyValuePair<bool, string>, object> _result = base.ObtenerCruces();
            if(_result.Key.Key) {
                string _JsonData = (string)_result.Value;
                File.WriteAllText($"{_Path}\\jsons\\Cruces.json", _JsonData);
                return new KeyValuePair<bool, string>(true, "Se generó correctamente el archivo json para los cruces.");
            }
            return _result.Key;
        }

        new KeyValuePair<bool, string> ObtenerPlantillas() {
            KeyValuePair<KeyValuePair<bool, string>, object> _result = base.ObtenerPlantillas();
            if(_result.Key.Key) {
                string _JsonData = (string)_result.Value, _FullPath = $"{_Path}\\jsons\\Plantillas.json";
                File.WriteAllText(_FullPath, _JsonData);
                oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>(_FullPath);
                foreach(oPlantilla _Template in _Templates) {
                    KeyValuePair<KeyValuePair<bool, string>, object> _resultFile = base.ObtenerArchivoPlantilla(_Template.IdPlantilla);
                    if(_resultFile.Key.Key) {
                        byte[] _TemplateFile = (byte[])_resultFile.Value;
                        File.WriteAllBytes($"{_Path}\\templates\\{_Template.Nombre}", _TemplateFile);
                        return new KeyValuePair<bool, string>(true, $"Se generó correctamente el archivo de la plantilla {_Template.Nombre}.");
                    }
                }
                return new KeyValuePair<bool, string>(true, "Se generó correctamente el archivo json para las plantillas.");
            }
            return _result.Key;
        }
    }
}