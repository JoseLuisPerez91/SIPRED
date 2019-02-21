using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Objects {
    public class oPlantilla : oBase {
        public oPlantilla(string _Usuario) : base(_Usuario) { }
        public int IdPlantilla { get; set; } = 0;
        public int IdTipoPlantilla { get; set; } = 0;
        public int Anio { get; set; } = 0;
        public string Nombre { get; set; } = "";
        public byte[] Plantilla { get; set; } = new byte[] { };
    }
}