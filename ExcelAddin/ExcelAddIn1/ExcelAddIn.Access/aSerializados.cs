using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace ExcelAddIn.Access {
    public class aSerializados : Connection {
        public aSerializados() { }
        /// <summary>Función para obtener el archivo Json de Cruces.
        /// <para>Invocar el SP dbo.spObtenerCruces de Tipo Scalar. Referencia: <see cref="base.ObtenerCruces()"/> se agrega la referencia ExcelAddIn.Access para invocarla.</para>
        /// <seealso cref="base.ObtenerCruces()"/>
        /// </summary>
        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerCruces() {
            return ExecuteScalar("[dbo].[spObtenerCruces]");
        }
        /// <summary>Función para obtener el archivo Json de Comprobaciones.
        /// <para>Invocar el SP dbo.spObtenerComprobaciones de Tipo Scalar. Referencia: <see cref="base.ObtenerComprobaciones()"/> se agrega la referencia ExcelAddIn.Access para invocarla.</para>
        /// <seealso cref="base.ObtenerComprobaciones()"/>
        /// </summary>
        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerComprobaciones() {
            return ExecuteScalar("[dbo].[spObtenerComprobaciones]");
        }
        /// <summary>Función para obtener el archivo Json de Tipos de Plantillas.
        /// <para>Invocar el SP dbo.spObtenerTiposPlantillas de Tipo Scalar. Referencia: <see cref="base.ObtenerTiposPlantillas()"/> se agrega la referencia ExcelAddIn.Access para invocarla.</para>
        /// <seealso cref="base.spObtenerTiposPlantillas()"/>
        /// </summary>
        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerTiposPlantillas() {
            return ExecuteScalar("[dbo].[spObtenerTiposPlantillas]");
        }
        /// <summary>Función para obtener el archivo Json de Plantillas.
        /// <para>Invocar el SP dbo.spObtenerPlantillas de Tipo Scalar. Referencia: <see cref="base.ObtenerPlantillas()"/> se agrega la referencia ExcelAddIn.Access para invocarla.</para>
        /// <seealso cref="base.ObtenerPlantillas()"/>
        /// </summary>
        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerPlantillas() {
            return ExecuteScalar("[dbo].[spObtenerPlantillas]");
        }
        /// <summary>Función para obtener el archivo Json de las validaciones de Cruces.
        /// <para>Invocar el SP dbo.spObtenerValidacionCruces de Tipo Scalar. Referencia: <see cref="base.ObtenerValidacionCruces()"/> se agrega la referencia ExcelAddIn.Access para invocarla.</para>
        /// <seealso cref="base.ObtenerValidacionCruces()"/>
        /// </summary>
        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerValidacionCruces()
        {
            return ExecuteScalar("[dbo].[spObtenerValidacionCruces]");
        }
        /// <summary>Función para obtener el archivo Json de Indices.
        /// <para>Invocar el SP dbo.spObtenerIndices de Tipo Scalar. Referencia: <see cref="base.ObtenerIndices()"/> se agrega la referencia ExcelAddIn.Access para invocarla.</para>
        /// <seealso cref="base.ObtenerIndices()"/>
        /// </summary>
        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerIndices()
        {
            return ExecuteScalar("[dbo].[spObtenerIndices]");
        }
        /// <summary>Función para obtener el archivo Json de la Plantilla.
        /// <para>Invocar el SP dbo.spObtenerArchivoPlantilla de Tipo Scalar. Referencia: <see cref="base.ObtenerArchivoPlantilla()"/> se agrega la referencia ExcelAddIn.Access para invocarla.</para>
        /// <seealso cref="base.ObtenerArchivoPlantilla()"/>
        /// </summary>
        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerArchivoPlantilla(int _IdPlantilla) {
            SqlParameter[] _Parameters = new SqlParameter[] { new SqlParameter("@pIdPlantilla", _IdPlantilla) };
            return ExecuteScalar("[dbo].[spObtenerArchivoPlantilla]", _Parameters);
        }
        /// <summary>Función para obtener la última versión de los archivos Json's.
        /// <para>Invocar el SP dbo.spObtenerTiposPlantillas de Tipo Scalar. Referencia: <see cref="base.ObtenerUpdate()"/> se agrega la referencia ExcelAddIn.Access para invocarla.</para>
        /// <seealso cref="base.ObtenerUpdate()"/>
        /// </summary>
        protected KeyValuePair<KeyValuePair<bool, string>, System.Data.DataTable> ObtenerUpdate()
        {
            return ExecuteTable("[dbo].[spObtenerIdTiposPlantillas]");
        }
        /// <summary>Función para obtener la última versión de los archivos Json's.
        /// <para>Invocar el SP dbo.spFormulasCMasivas de Tipo DataTable. Referencia: <see cref="base.ObtenerCMasiva()"/> se agrega la referencia ExcelAddIn.Access para invocarla.</para>
        /// <seealso cref="base.ObtenerCMasiva()"/>
        /// </summary>
        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerCMasiva()
        {
            return ExecuteScalar("[dbo].[spFormulasCMasivas]");
        }
    }
}