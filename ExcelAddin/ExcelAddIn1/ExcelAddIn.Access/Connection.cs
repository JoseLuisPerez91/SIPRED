using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace ExcelAddIn.Access {
    public class Connection {
        string _ConnectionString => string.Format(Configuration.ConnectionString, Configuration.Server, Configuration.DataBase, Configuration.User, Configuration.Password);
        public Connection() { }

        internal KeyValuePair<bool, string> ExecuteSP(string _Store, params SqlParameter[] _Parameters) {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "Se proceso correctamente la información");
            using(SqlConnection _Cnx = new SqlConnection(_ConnectionString)) {
                try {
                    _Cnx.Open();
                    using(SqlCommand _Cmd = new SqlCommand(_Store, _Cnx)) {
                        try {
                            if(_Parameters.Length > 0) _Cmd.Parameters.AddRange(_Parameters);
                            _Cmd.CommandType = System.Data.CommandType.StoredProcedure;
                            _Cmd.CommandTimeout = Configuration.TimeOut;
                            _Cmd.ExecuteNonQuery();
                        } catch(SqlException _sqlCmdEx) {
                            _result = new KeyValuePair<bool, string>(false, _sqlCmdEx.InnerException?.Message ?? _sqlCmdEx.Message);
                        } catch(Exception _cmdEx) {
                            _result = new KeyValuePair<bool, string>(false, _cmdEx.InnerException?.Message ?? _cmdEx.Message);
                        }
                    }
                } catch(SqlException _sqlEx) {
                    _result = new KeyValuePair<bool, string>(false, _sqlEx.InnerException?.Message ?? _sqlEx.Message);
                } catch(Exception _ex) {
                    _result = new KeyValuePair<bool, string>(false, _ex.InnerException?.Message ?? _ex.Message);
                } finally {
                    if(_Cnx.State == System.Data.ConnectionState.Open || _Cnx.State == System.Data.ConnectionState.Broken) _Cnx.Close();
                }
            }
            return _result;
        }
    }
}
