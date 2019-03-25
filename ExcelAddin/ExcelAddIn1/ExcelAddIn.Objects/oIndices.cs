using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Objects
{
    public class oIndices
    {
        public string Archivo { get; set; }
        public string Anexo { get;set; }
        public string Celda { get; set; }
        public int Cantidad { get; set; }
        public string Column { get; set; }
        public string Row { get; set; }
    }
}
