using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAddIn.Objects;
using ExcelAddIn.Access;

namespace ExcelAddIn.Logic {
    public class lPlantilla : aPlantilla
    {
        /// <summary>Función para generar el Archivo Excel según el Tipo de Plantilla.
        /// <para>Ejecuta la creación del Archivo Excel para el Tipo de Plantilla. Referencia: <see cref="lPlantilla(oPlantilla)"/> se agrega la referencia ExcelAddIn.Logic para invocarla.</para>
        /// <seealso cref="lPlantilla(oPlantilla)"/>
        /// </summary>
        public lPlantilla(oPlantilla _Template) : base(_Template) {
            Template = _Template;
            if(Template.IdTipoPlantilla == 0) _Messages.Add("Debe seleccionar un tipo.");
            if(Template.Anio == 0) _Messages.Add("Debe seleccionar un año.");
            if(string.IsNullOrEmpty(Template.Nombre) || string.IsNullOrWhiteSpace(Template.Nombre)) _Messages.Add("Debe seleccionar un archivo.");
            if(Template.Nombre.Length > 0 && Template.Plantilla.Length == 0) _Messages.Add("Debe seleccionar un archivo.");
        }
        /// <summary>Función para agregar los Mensajes al crear el Archivo Excel.
        /// <para>Agrega los Mensajes al crear el Archivo Excel. Referencia: <see cref="Add()"/> se agrega la referencia ExcelAddIn.Logic para invocarla.</para>
        /// <seealso cref="Add()"/>
        /// </summary>
        public new KeyValuePair<bool, string[]> Add() {
            if(_Messages.Count() > 0) return new KeyValuePair<bool, string[]>(false, _Messages.ToArray());
            KeyValuePair<KeyValuePair<bool, string>, int> _result = base.Add();
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }
    }
}