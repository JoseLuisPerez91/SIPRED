using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ExcelAddIn.Objects {
    public class oComprobacion {
        Regex regex = new Regex(@"\[.*?\]");

        public oComprobacion() { }

        public int IdComprobacion { get; set; }
        public int IdTipoPlantilla { get; set; }
        public string Concepto { get; set; }
        public string Formula { get; set; }
        public string Condicion { get; set; }
        public oCelda Destino { get; set; }
        public oCelda[] Celdas { get; set; }
        public oCelda[] CeldasCondicion { get; set; }
        public string FormulaExcel { get; set; }
        public string  Nota { get; set; }
        public int AdmiteCambios { get; set; }
        public bool EsFormula() => Celdas.Count() > 0 || CeldasCondicion.Count() > 0;

        public bool EsValida() => Celdas.Where(o => o.Fila == -1).Count() == 0 && CeldasCondicion.Where(o => o.Fila == -1).Count() == 0;

        public void setCeldas() {
            List<oCelda> _cells = new List<oCelda>();
            List<oCelda> _cCells = new List<oCelda>();
            Destino = new oCelda(Formula.Split('=')[0]);
            MatchCollection _match = regex.Matches(Formula.Split('=')[1]);
            MatchCollection _mCondicion = regex.Matches(Condicion);
            foreach(var _m in _match) _cells.Add(new oCelda(_m.ToString()));
            foreach(var _m in _mCondicion) _cCells.Add(new oCelda(_m.ToString()));
            Celdas = _cells.ToArray();
            CeldasCondicion = _cCells.ToArray();
        }

        public void setFormulaExcel() {
            if(EsValida() && EsFormula()) FormulaExcel = (CeldasCondicion.Count() > 0) ? $"IF({CeldasCondicion.ToString(Condicion, true)},{Celdas.ToString(Formula)},0)" : Celdas.ToString(Formula);
            if(EsValida() && !EsFormula()) FormulaExcel = Formula.Split('=')[1];
        }
    }
}