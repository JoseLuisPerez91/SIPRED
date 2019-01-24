using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ExcelAddIn.Objects {
    public class oCruce {
        Regex regex = new Regex(@"\[.*?\]");
        public oCruce() { }

        public int IdCruce { get; set; }
        public int IdTipoPlantilla { get; set; }
        public string Concepto { get; set; }
        public string Formula { get; set; }
        public string Condicion { get; set; }
        public oCelda[] CeldasFormula { get; private set; }
        public oCelda[] CeldasCondicion { get; private set; }
        public string FormulaExcel { get; private set; }
        public string CondicionExcel { get; private set; }
        public string ResultadoFormula { get; set; }
        public string ResultadoCondicion { get; set; }

        public void setCeldas() {
            List<oCelda> _cFormulas = new List<oCelda>();
            List<oCelda> _cCondicion = new List<oCelda>();
            MatchCollection _matchCF = regex.Matches(Formula);
            MatchCollection _matchCC = regex.Matches(Condicion);
            foreach(var _m in _matchCF) _cFormulas.Add(new oCelda(_m.ToString()));
            foreach(var _m in _matchCC) _cCondicion.Add(new oCelda(_m.ToString()));
            CeldasFormula = _cFormulas.ToArray();
            CeldasCondicion = _cCondicion.ToArray();
        }

        public void setFormulaExcel() {
            FormulaExcel = CeldasFormula.ToString(Formula, true);
            CondicionExcel = (CeldasCondicion.Count() > 0 && CeldasCondicion.Where(o => o.Fila == -1).Count() == 0) ? CeldasCondicion.ToString(Condicion, true) : "";
        }
    }
}