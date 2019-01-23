using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using ExcelAddIn.Objects;

namespace ExcelAddIn.Access {
    public class aPlantilla : Connection {
        protected oPlantilla Template = new oPlantilla("");
        public aPlantilla(oPlantilla _Template) : base() { }

        protected KeyValuePair<KeyValuePair<bool, string>, int> Add() {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "Se procesó corectamente la información.");
            SqlParameter[] _Parameters = {
                new SqlParameter("@pAnio", Template.Anio),
                new SqlParameter("@pIdTipoPlantilla",Template.IdTipoPlantilla),
                new SqlParameter("@pNombre", Template.Nombre),
                new SqlParameter("@pPlantilla", Template.Plantilla),
                new SqlParameter("@pUsuario", Template.Usuario)
            };
            return ExecuteNonQuery("[dbo].[spLoadTemplate]", _Parameters);
        }
    }
}