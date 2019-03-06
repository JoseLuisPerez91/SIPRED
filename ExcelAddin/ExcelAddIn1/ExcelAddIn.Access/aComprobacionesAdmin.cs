using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using ExcelAddIn.Objects;

namespace ExcelAddIn.Access
{
    public class aComprobacionesAdmin : Connection
    {
        protected oComprobacion Template = new oComprobacion();
        public aComprobacionesAdmin(oComprobacion _Template) : base() { }

        protected KeyValuePair<KeyValuePair<bool, string>, int> Add()
        {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "Se procesó corectamente la información.");
            SqlParameter[] _Parameters = {
                new SqlParameter("@pIdComprobacion", Template.IdComprobacion),
                new SqlParameter("@pIdTipoPlantilla",Template.IdTipoPlantilla),
                new SqlParameter("@pConcepto", Template.Concepto),
                new SqlParameter("@pFormula", Template.Formula),
                new SqlParameter("@pCondicion", Template.Condicion),
                new SqlParameter("@pNota", Template.Nota),
                new SqlParameter("@pAccion", "I")
            };
            return ExecuteNonQuery("[dbo].[spActualizarComprobaciones]", _Parameters);
        }

        protected KeyValuePair<KeyValuePair<bool, string>, int> Update()
        {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "Se procesó corectamente la información.");
            SqlParameter[] _Parameters = {
               new SqlParameter("@pIdComprobacion", Template.IdComprobacion),
                new SqlParameter("@pIdTipoPlantilla",Template.IdTipoPlantilla),
                new SqlParameter("@pConcepto", Template.Concepto),
                new SqlParameter("@pFormula", Template.Formula),
                new SqlParameter("@pCondicion", Template.Condicion),
                new SqlParameter("@pNota", Template.Nota),
                new SqlParameter("@pAccion", "M")
            };
            return ExecuteNonQuery("[dbo].[spActualizarComprobaciones]", _Parameters);
        }

        protected KeyValuePair<KeyValuePair<bool, string>, int> Delete()
        {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "Se procesó corectamente la información.");
            SqlParameter[] _Parameters = {
             new SqlParameter("@pIdComprobacion", Template.IdComprobacion),
                new SqlParameter("@pIdTipoPlantilla",Template.IdTipoPlantilla),
                new SqlParameter("@pConcepto", Template.Concepto),
                new SqlParameter("@pFormula", Template.Formula),
                new SqlParameter("@pCondicion", Template.Condicion),
                new SqlParameter("@pNota", Template.Nota),
                new SqlParameter("@pAccion", "E")
            };
            return ExecuteNonQuery("[dbo].[spActualizarComprobaciones]", _Parameters);
        }
    }
}
