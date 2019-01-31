using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;

namespace ExcelAddIn.Objects {
    public static class Assembler {
        public static T LoadJson<T>(string _Path) => JsonConvert.DeserializeObject<T>(File.ReadAllText(_Path));

        public static string ToString(this oCelda[] _Cells, string _Formula, bool _Condicion = false) {
            string _result = (!_Condicion) ? _Formula.Split('=')[1] : _Formula;
            foreach(oCelda _cell in _Cells) _result = _result.Replace(_cell.Original, _cell.CeldaExcel);
            return _result;
        }

        public static string ToString(this oCeldaCondicion[] _Cells, string _Formula, bool _Condicion = false)
        {
            string _result = (!_Condicion) ? _Formula.Split('=')[1] : _Formula;
            foreach (oCeldaCondicion _cell in _Cells) _result = _result.Replace(_cell.Original, _cell.CeldaExcel);
            return _result;
        }
    }
}
