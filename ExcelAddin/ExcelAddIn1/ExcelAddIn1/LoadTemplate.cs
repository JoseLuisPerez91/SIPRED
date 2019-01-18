using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using Newtonsoft.Json;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;

namespace ExcelAddIn1 {
    public partial class LoadTemplate : Base {
        public LoadTemplate() {
            InitializeComponent();
            FillTemplateType(cmbTipoPlantilla);
            FillYears(cmbAnio);
        }

        private void btnSeleccionar_Click(object sender, EventArgs e) {
            DialogResult _Result = ofdTemplate.ShowDialog();
        }

        private void ofdTemplate_FileOk(object sender, CancelEventArgs e) { txtPlantilla.Text = ofdTemplate.FileName; }

        private void btnCancelar_Click(object sender, EventArgs e) {
            cmbAnio.SelectedIndex = 0;
            cmbTipoPlantilla.SelectedIndex = 0;
            txtPlantilla.Text = "";
        }

        private void btnCargar_Click(object sender, EventArgs e) {
            oPlantilla _Template = new oPlantilla("eduardo.perez") {
                Anio = (int)cmbAnio.SelectedValue,
                IdTipoPlantilla = (int)cmbTipoPlantilla.SelectedValue,
                Nombre = new FileInfo(txtPlantilla.Text).Name,
                Plantilla = File.ReadAllBytes(txtPlantilla.Text)
            };
            KeyValuePair<bool, string[]> _result = new lPlantilla(_Template).Add();
            string _Messages = "";
            foreach(string _Msg in _result.Value) _Messages += ((_Messages.Length > 0) ? "\r\n" : "") + _Msg;
            MessageBox.Show(_Messages, (_result.Key) ? "Proceso Existoso" : "Información Faltante", MessageBoxButtons.OK, (_result.Key) ? MessageBoxIcon.Information : MessageBoxIcon.Exclamation);
            if(_result.Key) btnCancelar_Click(btnCancelar, null);
        }
    }
}