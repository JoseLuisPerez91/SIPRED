using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ExcelAddIn.Objects {
    public class oCelda {
        public oCelda(string _Expression) {
            Original = _Expression;
            Anexo = $"Anexo {int.Parse(_Expression.Replace("[", "").Replace("]", "").Split(',')[0])}";
            Indice = _Expression.Replace("[", "").Replace("]", "").Split(',')[1];
            Columna = int.Parse(_Expression.Replace("[", "").Replace("]", "").Split(',')[2]);
        }

        public string Original { get; private set; }
        public string Anexo { get; set; }
        public string Indice { get; set; }
        public int Columna { get; set; }
        public int Fila { get; set; } = -1;
        public string CeldaExcel { get; private set; }

        public void setCeldaExcel(ExcelRange _Cell, string _Anexo) {
            Anexo = _Cell.Worksheet.Name;
            CeldaExcel = (Anexo == _Anexo || _Anexo == "") ? _Cell.Address : _Cell.FullAddress;
        }
    }
}
