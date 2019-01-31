using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using System.Xml;

namespace ExcelAddIn.Access {
    internal enum ExecutionType {
        DataSet,
        DataTable,
        NonQuery,
        Scalar,
        Reader,
        XmlReader
    }

    public class Connection {
        string _ConnectionString => string.Format(Configuration.ConnectionString, Configuration.Server, Configuration.DataBase, Configuration.User, Configuration.Password);

        protected List<string> _Messages = new List<string>();

        public Connection() { }

        object Execute(string _Store, ExecutionType _Type, params SqlParameter[] _Parameters) {
            object _result = null;
            using(SqlConnection _Cnx = new SqlConnection(_ConnectionString)) {
                try {
                    _Cnx.Open();
                    using(SqlCommand _Cmd = new SqlCommand(_Store, _Cnx)) {
                        try {
                            if(_Parameters.Length > 0) _Cmd.Parameters.AddRange(_Parameters);
                            _Cmd.CommandType = System.Data.CommandType.StoredProcedure;
                            _Cmd.CommandTimeout = Configuration.TimeOut;
                            switch(_Type) {
                                case ExecutionType.DataSet | ExecutionType.DataTable:
                                    using(SqlDataAdapter _sqlDA = new SqlDataAdapter(_Cmd)) {
                                        if(_Type == ExecutionType.DataSet) {
                                            DataSet _ds = new DataSet();
                                            _sqlDA.Fill(_ds);
                                            _result = _ds;
                                        } else if(_Type == ExecutionType.DataTable) {
                                            DataTable _dt = new DataTable();
                                            _sqlDA.Fill(_dt);
                                            _result = _dt;
                                        }
                                    }
                                    break;
                                case ExecutionType.NonQuery:
                                    _result = _Cmd.ExecuteNonQuery();
                                    break;
                                case ExecutionType.Reader:
                                    _result = _Cmd.ExecuteReader();
                                    break;
                                case ExecutionType.Scalar:
                                    _result = _Cmd.ExecuteScalar();
                                    break;
                                case ExecutionType.XmlReader:
                                    _result = _Cmd.ExecuteXmlReader();
                                    break;
                            }
                        } catch(SqlException _sqlCmdEx) {
                            throw _sqlCmdEx;
                        } catch(Exception _cmdEx) {
                            throw _cmdEx;
                        }
                    }
                } catch(SqlException _sqlEx) {
                    throw _sqlEx;
                } catch(Exception _ex) {
                    throw _ex;
                } finally {
                    if(_Cnx.State == ConnectionState.Open || _Cnx.State == ConnectionState.Broken) _Cnx.Close();
                }
            }
            return _result;
        }

        internal KeyValuePair<KeyValuePair<bool, string>, int> ExecuteNonQuery(string _Store, params SqlParameter[] _Parameters) {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "Se proceso correctamente la información");
            int _executeResult = 0;
            try {
                _executeResult = (int)Execute(_Store, ExecutionType.NonQuery, _Parameters);
            } catch(SqlException _sqlEx) {
                _result = new KeyValuePair<bool, string>(false, _sqlEx.InnerException?.Message ?? _sqlEx.Message);
            } catch(Exception _ex) {
                _result = new KeyValuePair<bool, string>(false, _ex.InnerException?.Message ?? _ex.Message);
            }
            return new KeyValuePair<KeyValuePair<bool, string>, int>(_result, _executeResult);
        }

        internal KeyValuePair<KeyValuePair<bool, string>, DataSet> ExecuteDataSet(string _Store, params SqlParameter[] _Parameters) {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "");
            DataSet _executeResult = null;
            try {
                _executeResult = (DataSet)Execute(_Store, ExecutionType.DataSet, _Parameters);
            } catch(SqlException _sqlEx) {
                _result = new KeyValuePair<bool, string>(false, _sqlEx.InnerException?.Message ?? _sqlEx.Message);
            } catch(Exception _ex) {
                _result = new KeyValuePair<bool, string>(false, _ex.InnerException?.Message ?? _ex.Message);
            }
            return new KeyValuePair<KeyValuePair<bool, string>, DataSet>(_result, _executeResult);
        }

        internal KeyValuePair<KeyValuePair<bool, string>, DataTable> ExecuteTable(string _Store, params SqlParameter[] _Parameters) {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "");
            DataTable _executeResult = null;
            try {
                _executeResult = (DataTable)Execute(_Store, ExecutionType.DataTable, _Parameters);
            } catch(SqlException _sqlEx) {
                _result = new KeyValuePair<bool, string>(false, _sqlEx.InnerException?.Message ?? _sqlEx.Message);
            } catch(Exception _ex) {
                _result = new KeyValuePair<bool, string>(false, _ex.InnerException?.Message ?? _ex.Message);
            }
            return new KeyValuePair<KeyValuePair<bool, string>, DataTable>(_result, _executeResult);
        }

        internal KeyValuePair<KeyValuePair<bool, string>, object> ExecuteScalar(string _Store, params SqlParameter[] _Parameters) {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "");
            object _executeResult = null;
            try {
                _executeResult = Execute(_Store, ExecutionType.Scalar, _Parameters);
            } catch(SqlException _sqlEx) {
                _result = new KeyValuePair<bool, string>(false, _sqlEx.InnerException?.Message ?? _sqlEx.Message);
            } catch(Exception _ex) {
                _result = new KeyValuePair<bool, string>(false, _ex.InnerException?.Message ?? _ex.Message);
            }
            return new KeyValuePair<KeyValuePair<bool, string>, object>(_result, _executeResult);
        }

        internal KeyValuePair<KeyValuePair<bool, string>, SqlDataReader> ExecuteReader(string _Store, params SqlParameter[] _Parameters) {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "");
            SqlDataReader _executeResult = null;
            try {
                _executeResult = (SqlDataReader)Execute(_Store, ExecutionType.NonQuery, _Parameters);
            } catch(SqlException _sqlEx) {
                _result = new KeyValuePair<bool, string>(false, _sqlEx.InnerException?.Message ?? _sqlEx.Message);
            } catch(Exception _ex) {
                _result = new KeyValuePair<bool, string>(false, _ex.InnerException?.Message ?? _ex.Message);
            }
            return new KeyValuePair<KeyValuePair<bool, string>, SqlDataReader>(_result, _executeResult);
        }

        internal KeyValuePair<KeyValuePair<bool, string>, XmlReader> ExecuteXmlReader(string _Store, params SqlParameter[] _Parameters) {
            KeyValuePair<bool, string> _result = new KeyValuePair<bool, string>(true, "");
            XmlReader _executeResult = null;
            try {
                _executeResult = (XmlReader)Execute(_Store, ExecutionType.NonQuery, _Parameters);
            } catch(SqlException _sqlEx) {
                _result = new KeyValuePair<bool, string>(false, _sqlEx.InnerException?.Message ?? _sqlEx.Message);
            } catch(Exception _ex) {
                _result = new KeyValuePair<bool, string>(false, _ex.InnerException?.Message ?? _ex.Message);
            }
            return new KeyValuePair<KeyValuePair<bool, string>, XmlReader>(_result, _executeResult);
        }
    }
}
