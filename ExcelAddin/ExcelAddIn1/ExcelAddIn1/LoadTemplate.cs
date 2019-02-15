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
using ExcelAddIn.Access;

namespace ExcelAddIn1
{
    public partial class LoadTemplate : Base
    {
        public LoadTemplate()
        {
            string _Path = Configuration.Path;
            InitializeComponent();
            
            if (Directory.Exists(_Path + "\\jsons") && Directory.Exists(_Path + "\\templates"))
            {
                if (File.Exists(_Path + "\\jsons\\TiposPlantillas.json"))
                {
                    FillTemplateType(cmbTipoPlantilla);
                    FillYears(cmbAnio);
                }
                else
                {
                    this.TopMost = false;
                    this.Enabled = false;
                    this.Hide();
                    FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                    _FileJsonfrm._Form = this;
                    _FileJsonfrm._Process = false;
                    _FileJsonfrm._window = this.Text;
                    _FileJsonfrm.Show();
                    return;
                }
            }
            else
            {
                if (!Directory.Exists(_Path + "\\jsons"))
                {
                    Directory.CreateDirectory(_Path + "\\jsons");
                }
                if (!Directory.Exists(_Path + "\\templates"))
                {
                    Directory.CreateDirectory(_Path + "\\templates");
                }

                this.TopMost = false;
                this.Enabled = false;
                this.Hide();
                FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                _FileJsonfrm._Form = this;
                _FileJsonfrm._Process = false;
                _FileJsonfrm._window = this.Text;
                _FileJsonfrm.Show();
                return;
            }
        }

        private void btnSeleccionar_Click(object sender, EventArgs e)
        {
            DialogResult _Result = ofdTemplate.ShowDialog();
        }

        private void ofdTemplate_FileOk(object sender, CancelEventArgs e) { txtPlantilla.Text = ofdTemplate.FileName; }

        private void btnCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnCargar_Click(object sender, EventArgs e)
        {
            string _Message = (cmbTipoPlantilla.SelectedIndex == 0) ? "- Debe seleccionar un tipo." : "";
            _Message += (cmbAnio.SelectedIndex == 0) ? ((_Message.Length > 0) ? "\r\n" : "") + "- Debe seleccionar un año." : "";
            _Message += (txtPlantilla.Text == "") ? ((_Message.Length > 0) ? "\r\n" : "") + "- Debe seleccionar un archivo." : "";
            if (_Message.Length > 0)
            {
                MessageBox.Show(_Message, "Información Faltante", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>($"{Configuration.Path}\\jsons\\Plantillas.json");
            oPlantilla _Template = new oPlantilla("app_sipred")
            {
                Anio = (int)cmbAnio.SelectedValue,
                IdTipoPlantilla = (int)cmbTipoPlantilla.SelectedValue,
                Nombre = new FileInfo(txtPlantilla.Text).Name,
                Plantilla = File.ReadAllBytes(txtPlantilla.Text)
            };
            DialogResult _response = DialogResult.None;
            if (_Templates.Where(o => o.IdTipoPlantilla == _Template.IdTipoPlantilla && o.Anio == _Template.Anio).Count() > 0)
            {
                _response = MessageBox.Show($"¿Desea reemplazar la plantilla para {((oTipoPlantilla)cmbTipoPlantilla.SelectedItem).FullName} y {cmbAnio.SelectedValue.ToString()}?", "Plantilla Existente", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (_response == DialogResult.No)
                {
                    btnCancelar_Click(btnCancelar, null);
                    return;
                }
            }
            KeyValuePair<bool, string[]> _result = new lPlantilla(_Template).Add();
            string _Messages = "";
            foreach (string _Msg in _result.Value) _Messages += ((_Messages.Length > 0) ? "\r\n" : "") + _Msg;
            if (_result.Key && _response != DialogResult.Yes) _Messages = "La plantilla fue reemplazada con éxito";
            MessageBox.Show(_Messages, (_result.Key) ? "Proceso Existoso" : "Información Faltante", MessageBoxButtons.OK, (_result.Key) ? MessageBoxIcon.Information : MessageBoxIcon.Exclamation);
            if (_result.Key) btnCancelar_Click(btnCancelar, null);
        }
    }
}