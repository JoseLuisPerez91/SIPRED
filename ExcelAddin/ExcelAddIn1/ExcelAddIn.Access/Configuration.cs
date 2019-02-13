using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace ExcelAddIn.Access {
    public static class Configuration {
        static string _getConfig(string _Key) => ConfigurationManager.AppSettings[_Key];
        static string _unEncrypt(string _Value) => Encoding.UTF8.GetString(Convert.FromBase64String(_Value));

        public static string ConnectionString => _unEncrypt(_getConfig("VAL0"));
        public static string Server => _unEncrypt(_getConfig("VAL1"));
        public static string DataBase => _unEncrypt(_getConfig("VAL2"));
        public static string User => _unEncrypt(_getConfig("VAL3"));
        public static string Password => _unEncrypt(_getConfig("VAL4"));
        public static int TimeOut => int.Parse(_unEncrypt(_getConfig("VAL5")));
        public static string Path => _getConfig("VAL6");
        public static int PwsExcel => int.Parse(_unEncrypt(_getConfig("VAL7")));
    }
}