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
using ExcelAddIn1.Objects;
using ExcelAddIn1.Assemblers;
using ExcelAddIn1.DataAccess;

namespace ExcelAddIn1 {
    public partial class LoadTemplate : Form {
        public LoadTemplate() {
            InitializeComponent();
            FillTemplateType();
            FillYears();
        }

        void FillYears() {
            DateTime _Now = DateTime.Now;
            oAnio[] _Years = { new oAnio() { Id = _Now.Year - 1, Concepto = (_Now.Year - 1).ToString() }, new oAnio() { Id = _Now.Year, Concepto = _Now.Year.ToString() } };
            cmbAnio.Fill<oAnio>(_Years, "Id", "Concepto", new oAnio() { Id = 0, Concepto = "Seleccione un Año" });
        }

        void FillTemplateType() {
            oTipoPlantilla[] _TemplatesTypes = Assembler.LoadJson<oTipoPlantilla>(Environment.CurrentDirectory + "jsons\\TemplatesTypes.json");
            cmbTipoPlantilla.Fill<oTipoPlantilla>(_TemplatesTypes, "Id", "FullName", new oTipoPlantilla() { IdTipoPlantilla = 0, Clave = "", Concepto = "Seleccione un Tipo de Plantilla" });
        }

        private void btnSearch_Click(object sender, EventArgs e) {
            DialogResult _Result = ofdTemplate.ShowDialog();
        }

        private void ofdTemplate_FileOk(object sender, CancelEventArgs e) {
            txtPlantilla.Text = ofdTemplate.FileName;
        }

        private void btnCancelar_Click(object sender, EventArgs e) {
            cmbAnio.SelectedValue = "0";
            cmbTipoPlantilla.SelectedValue = "0";
            txtPlantilla.Text = "";
        }

        private void btnCargar_Click(object sender, EventArgs e) {
            string _Message = (cmbAnio.SelectedValue.ToString() == "0") ? "Debe seleccionar un año." : "";
            _Message += (cmbTipoPlantilla.SelectedValue.ToString() == "0") ? ((_Message.Length > 0) ? "\r\n" : "") + "Debe seleccionar un tipo." : "";
            if(_Message.Length > 0) {
                MessageBox.Show(_Message, "Información Faltante", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            FileInfo _TemplateFile = new FileInfo(txtPlantilla.Text);
            SqlParameter[] _Parameters = {
                new SqlParameter("@pAnio", (int)cmbAnio.SelectedValue),
                new SqlParameter("@pIdTipoPlantilla", (int)cmbTipoPlantilla.SelectedValue),
                new SqlParameter("@pNombre", _TemplateFile.Name),
                new SqlParameter("@pPlantilla", File.ReadAllBytes(txtPlantilla.Text)),
                new SqlParameter("@pUsuario", "eduardo.perez")
            };
            Connection _Cnx = new Connection();
            _Cnx.ExecuteSP("[dbo].[spLoadTemplate]", _Parameters);
        }
    }
}
