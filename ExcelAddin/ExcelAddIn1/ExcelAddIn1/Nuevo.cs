using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExcelAddIn.Objects;
using ExcelAddIn1.Assemblers;

namespace ExcelAddIn1 {
    public partial class Nuevo : Base {
        public Nuevo() {
            InitializeComponent();
            FillYears(cmbAnio);
            FillTemplateType(cmbTipo);
        }

        private void btnCancelar_Click(object sender, EventArgs e) {
            cmbAnio.SelectedIndex = 0;
            cmbTipo.SelectedIndex = 0;
        }

        private void btnCrear_Click(object sender, EventArgs e) {
            oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>($"{Environment.CurrentDirectory}\\jsons\\Templates.json");
            int _IdTemplateType = (int)cmbTipo.SelectedValue, _Year = (int)cmbAnio.SelectedValue;
            oPlantilla _Template = _Templates.FirstOrDefault(o => o.IdTipoPlantilla == _IdTemplateType && o.Anio == _Year);
            if(_Template != null) {
                MessageBox.Show("No existe una plantilla para el tipo seleccionado, favor de seleccionar otro tipo o contactar al administrador.", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            string _newTemplate = $"{Environment.CurrentDirectory}\\{((oTipoPlantilla)cmbTipo.SelectedItem).Clave}-{cmbAnio.SelectedValue.ToString()}-{DateTime.Now.ToString("ddMMyyyyHHmmss")}.xlsx";
            string _currentTemplate = $"{Environment.CurrentDirectory}\\templates\\{_Template.Nombre}";
            Microsoft.Office.Interop.Excel.Application _current = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook _workbook = _current.Workbooks.Open(_currentTemplate, Microsoft.Office.Interop.Excel.XlUpdateLinks.xlUpdateLinksNever, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _workbook.SaveAs(_newTemplate, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel12, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            _workbook.Close(false, Type.Missing, Type.Missing);
            _current.Quit();
            Globals.ThisAddIn.Application.Visible = true;
            Globals.ThisAddIn.Application.Workbooks.Open(_newTemplate);
        }
    }
}