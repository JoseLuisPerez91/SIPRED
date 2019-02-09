using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn.Objects
{
    public class oRootobject
    {
        public List<oSubtotal> Subtotales { get; set; }
        public List<oConcepto> Conceptos { get; set; }

    }
}
