using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace ExcelAddIn.Access {
    public class aSerializados : Connection {
        KeyValuePair<KeyValuePair<bool, string>, object> GetJson(string _Store) {
            return ExecuteScalar(_Store);
        }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerCruces() {
            return GetJson("[dbo].[spObtenerCruces]");
        }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerComprobaciones() {
            return GetJson("[dbo].[spObtenerComprobaciones]");
        }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerTiposPlantillas() {
            return GetJson("[dbo].[spObtenerTiposPlantillas]");
        }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerPlantillas() {
            return GetJson("[dbo].[spObtenerPlantillas]");
        }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerArchivoPlantilla(int _IdPlantilla) {
            SqlParameter[] _Parameters = new SqlParameter[] { new SqlParameter("@pIdPlantilla", _IdPlantilla) };
            return ExecuteScalar("[dbo].[spObtenerArchivoPlantilla]", _Parameters);
        }
    }
}