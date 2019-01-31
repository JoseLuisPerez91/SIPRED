using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace ExcelAddIn.Access {
    public class aSerializados : Connection {
        public aSerializados() { }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerCruces() {
            return ExecuteScalar("[dbo].[spObtenerCruces]");
        }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerComprobaciones() {
            return ExecuteScalar("[dbo].[spObtenerComprobaciones]");
        }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerTiposPlantillas() {
            return ExecuteScalar("[dbo].[spObtenerTiposPlantillas]");
        }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerPlantillas() {
            return ExecuteScalar("[dbo].[spObtenerPlantillas]");
        }

        protected KeyValuePair<KeyValuePair<bool, string>, object> ObtenerArchivoPlantilla(int _IdPlantilla) {
            SqlParameter[] _Parameters = new SqlParameter[] { new SqlParameter("@pIdPlantilla", _IdPlantilla) };
            return ExecuteScalar("[dbo].[spObtenerArchivoPlantilla]", _Parameters);
        }
    }
}