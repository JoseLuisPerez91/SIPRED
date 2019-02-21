using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;

namespace ExcelAddIn.Objects {
    public static class Assembler
    {
        /// <summary>Función para cargar los archivos Json.
        /// <para>Carga el archivo Json según la ruta especificada. Referencia: <see cref="LoadJson{T}(string)"/> se agrega la referencia ExcelAddIn.Object para invocarla.</para>
        /// <seealso cref="LoadJson{T}(string)"/>
        /// </summary>
        public static T LoadJson<T>(string _Path) => JsonConvert.DeserializeObject<T>(File.ReadAllText(_Path));
        public static string ToString(this oCelda[] _Cells, string _Formula, bool _Condicion = false) {
            string _result = (!_Condicion) ? _Formula.Split('=')[1] : _Formula;
            foreach(oCelda _cell in _Cells) _result = _result.Replace(_cell.Original, _cell.CeldaExcel);
            return _result;
        }
        /// <summary>Función de Tipo String.
        /// <para>Función para convertir a tipo String las Celdas, Formulas y Condición del Archivo Json. Referencia: <see cref="ToString(oCeldaCondicion[], string, bool)"/> se agrega la referencia ExcelAddIn.Object para invocarla.</para>
        /// <seealso cref="ToString(oCeldaCondicion[], string, bool)"/>
        /// </summary>
        public static string ToString(this oCeldaCondicion[] _Cells, string _Formula, bool _Condicion = false)
        {
            string _result = (!_Condicion) ? _Formula.Split('=')[1] : _Formula;
            foreach (oCeldaCondicion _cell in _Cells) _result = _result.Replace(_cell.Original, _cell.CeldaExcel);
            return _result;
        }
    }
}
