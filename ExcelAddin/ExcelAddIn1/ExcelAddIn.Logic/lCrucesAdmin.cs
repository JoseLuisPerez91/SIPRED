using ExcelAddIn.Access;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelAddIn.Objects;


namespace ExcelAddIn.Logic
{
    public class lCrucesAdmin: aCrucesAdmin
    {
        public lCrucesAdmin(oCruce _Template,string accion) : base(_Template)
        {
            Template = _Template;
            if ((accion == "A") || (accion == "M"))
            {
                if (Template.IdTipoPlantilla == 0) _Messages.Add("Debe seleccionar un tipo.");
                if (string.IsNullOrEmpty(Template.Concepto) || string.IsNullOrWhiteSpace(Template.Concepto)) _Messages.Add("Debe seleccionar un concepto.");
                if (string.IsNullOrEmpty(Template.Formula) || string.IsNullOrWhiteSpace(Template.Formula)) _Messages.Add("Debe seleccionar un formula.");
            }

        }

        public new KeyValuePair<bool, string[]> Add()
        {
            if (_Messages.Count() > 0) return new KeyValuePair<bool, string[]>(false, _Messages.ToArray());
            KeyValuePair<KeyValuePair<bool, string>, int> _result = base.Add();
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }

        public new KeyValuePair<bool, string[]> Update()
        {
            if (_Messages.Count() > 0) return new KeyValuePair<bool, string[]>(false, _Messages.ToArray());
            KeyValuePair<KeyValuePair<bool, string>, int> _result = base.Update();
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }


        public new KeyValuePair<bool, string[]> Delete()
        {
            if (_Messages.Count() > 0) return new KeyValuePair<bool, string[]>(false, _Messages.ToArray());
            KeyValuePair<KeyValuePair<bool, string>, int> _result = base.Delete();
            return new KeyValuePair<bool, string[]>(_result.Key.Key, new string[] { _result.Key.Value });
        }
    }
}
