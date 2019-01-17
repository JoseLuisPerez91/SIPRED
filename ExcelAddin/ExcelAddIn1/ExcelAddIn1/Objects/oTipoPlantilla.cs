using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn1.Objects {
    public class oTipoPlantilla {
        public oTipoPlantilla() { }

        public int IdTipoPlantilla { get; set; }
        public string Clave { get; set; }
        public string Concepto { get; set; }
        public string FullName => $"{((Clave.Length > 0) ? $"{Clave} - " : "")}{Concepto}";
    }
}
