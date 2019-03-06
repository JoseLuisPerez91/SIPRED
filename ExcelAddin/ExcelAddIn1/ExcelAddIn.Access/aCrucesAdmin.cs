using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using ExcelAddIn.Objects;

namespace ExcelAddIn.Access
{
    public class aCrucesAdmin:Connection
    {
        protected oCruce Template = new oCruce();
        public aCrucesAdmin(oCruce _Template) : base() { }

        protected KeyValuePair<KeyValuePair<bool, string>, int> Add()
        {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "Se procesó corectamente la información.");
            SqlParameter[] _Parameters = {
                new SqlParameter("@pIdCruce", Template.IdCruce),
                new SqlParameter("@pIdTipoPlantilla",Template.IdTipoPlantilla),
                new SqlParameter("@pConcepto", Template.Concepto),
                new SqlParameter("@pFormula", Template.Formula),
                new SqlParameter("@pCondicion", Template.Condicion),
                new SqlParameter("@pNota", Template.Nota),
                new SqlParameter("@pLecImportes", Template.LecturaImportes),
                new SqlParameter("@pAccion", "I")
            };
            return ExecuteNonQuery("[dbo].[spActualizarCruces]", _Parameters);
        }

        protected KeyValuePair<KeyValuePair<bool, string>, int> Update()
        {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "Se procesó corectamente la información.");
            SqlParameter[] _Parameters = {
                new SqlParameter("@pIdCruce", Template.IdCruce),
                new SqlParameter("@pIdTipoPlantilla",Template.IdTipoPlantilla),
                new SqlParameter("@pConcepto", Template.Concepto),
                new SqlParameter("@pFormula", Template.Formula),
                new SqlParameter("@pCondicion", Template.Condicion),
                new SqlParameter("@pNota", Template.Nota),
                new SqlParameter("@pLecImportes", Template.LecturaImportes),
                new SqlParameter("@pAccion", "M")
            };
            return ExecuteNonQuery("[dbo].[spActualizarCruces]", _Parameters);
        }

        protected KeyValuePair<KeyValuePair<bool, string>, int> Delete()
        {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "Se procesó corectamente la información.");
            SqlParameter[] _Parameters = {
                new SqlParameter("@pIdCruce", Template.IdCruce),
                new SqlParameter("@pIdTipoPlantilla",Template.IdTipoPlantilla),
                new SqlParameter("@pConcepto", Template.Concepto),
                new SqlParameter("@pFormula", Template.Formula),
                new SqlParameter("@pCondicion", Template.Condicion),
                new SqlParameter("@pNota", Template.Nota),
                new SqlParameter("@pLecImportes", Template.LecturaImportes),
                new SqlParameter("@pAccion", "E")
            };
            return ExecuteNonQuery("[dbo].[spActualizarCruces]", _Parameters);
        }
    }
}
