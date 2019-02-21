using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Objects {
    public class oTipoPlantilla {
        public oTipoPlantilla() { }
        public int IdTipoPlantilla { get; set; } = 0;
        public string Clave { get; set; } = "";
        public string Concepto { get; set; } = "";
        public string FullName => $"{((Clave.Length > 0) ? $"{Clave} - " : "")}{Concepto}";
    }
}