using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Objects {
    public class oBase {
        public oBase(string _Usuario) { Usuario = _Usuario; }
        public string Usuario { get; set; } = "";
    }
}