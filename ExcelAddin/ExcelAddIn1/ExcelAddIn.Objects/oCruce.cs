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
        public oCelda[] CeldasFormula { get; set; }
        public oCeldaCondicion[] CeldasCondicion { get; set; }
        public string FormulaExcel { get; set; }
        public string CondicionExcel { get; set; }
        public string ResultadoFormula { get; set; }
        public string ResultadoCondicion { get; set; }
        public string Diferencia { get; set; }
        public string Grupo1 { get; set; }
        public string Grupo2 { get; set; }
        public int LecturaImportes { get; set; }
        public string Nota { get; set; }

        public void setCeldas() {
            List<oCelda> _cFormulas = new List<oCelda>();
            List<oCeldaCondicion> _cCondicion = new List<oCeldaCondicion>();
            MatchCollection _matchCF = regex.Matches(Formula);
            MatchCollection _matchCC = regex.Matches(Condicion);
            foreach(var _m in _matchCF) _cFormulas.Add(new oCelda(_m.ToString()));
            foreach(var _m in _matchCC) _cCondicion.Add(new oCeldaCondicion(_m.ToString()));
            CeldasFormula = _cFormulas.ToArray();
            CeldasCondicion = _cCondicion.ToArray();
        }

        public void setFormulaExcel() {
            FormulaExcel = CeldasFormula.ToString(Formula, true);
            CondicionExcel = (CeldasCondicion.Count() > 0 && CeldasCondicion.Where(o => o.Fila == -1).Count() == 0) ? CeldasCondicion.ToString(Condicion, true) : "";
        }
    }
}