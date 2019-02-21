using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;

namespace ExcelAddIn.Objects {
    public class oValidaCruces
    {
        public oValidaCruces() { }
        public string  Hoja { get; set; }
        public string Indice { get; set; }
        public string Concepto { get; set; } = "";
        public bool EsCorrecto { get; set; } = true;

    }
}