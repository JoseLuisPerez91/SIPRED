using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;

namespace ExcelAddIn1.Assemblers {
    public static class Assembler {
        public static void Fill<T>(this ComboBox _cmb, T[] _Source, string _ValueField, string _TextField, T _Initial) {
            List<T> _FinalSource = new List<T>();
            if(_Initial != null)
                _FinalSource.Add(_Initial);
            _FinalSource.AddRange(_Source);
            _cmb.Items.Clear();
            _cmb.DisplayMember = _TextField;
            _cmb.ValueMember = _ValueField;
            _cmb.DataSource = _FinalSource;
        }

        public static T[] LoadJson<T>(string _Path) {
            string _Content = File.ReadAllText(_Path);
            return JsonConvert.DeserializeObject<T[]>(_Content);
        }
    }
}
