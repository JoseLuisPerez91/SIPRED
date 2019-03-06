using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAddIn.Objects;
using ExcelAddIn1.Assemblers;

namespace ExcelAddIn1 {
    public class Base : Form {
        public Base() { }
        internal void FillYears(ComboBox _cmb) {
            DateTime _Now = DateTime.Now;
            oAnio[] _Years = { new oAnio() { Id = _Now.Year - 1, Concepto = (_Now.Year - 1).ToString() }, new oAnio() { Id = _Now.Year, Concepto = _Now.Year.ToString() } };
            _cmb.Fill<oAnio>(_Years, "Id", "Concepto", new oAnio() { Id = 0, Concepto = "Seleccione un Año" });
        }
        internal void FillTemplateType(ComboBox _cmb) {
            oTipoPlantilla[] _TemplatesTypes = ExcelAddIn.Objects.Assembler.LoadJson<oTipoPlantilla[]>($"{ExcelAddIn.Access.Configuration.Path}\\jsons\\TiposPlantillas.json");
            _cmb.Fill<oTipoPlantilla>(_TemplatesTypes, "IdTipoPlantilla", "FullName", new oTipoPlantilla() { IdTipoPlantilla = 0, Clave = "", Concepto = "Seleccione un Tipo de Plantilla" });
        }
        internal static DataTable ToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection props =
            TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, prop.PropertyType);

            }
            object[] values = new object[props.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item);
                }
                table.Rows.Add(values);
            }
            return table;
        }
    }
}