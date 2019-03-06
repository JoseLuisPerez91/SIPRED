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
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            InitializeComponent();
            
            if (Directory.Exists(_Path + "\\jsons") && Directory.Exists(_Path + "\\templates"))
            {
                if (File.Exists(_Path + "\\jsons\\TiposPlantillas.json"))
                {
                    if (_Connection)
                    {
                        KeyValuePair<bool, System.Data.DataTable> _TipoPlantilla = new lSerializados().ObtenerUpdate();

                        foreach (DataRow _Row in _TipoPlantilla.Value.Rows)
                        {
                            string _IdTipoPlantilla = _Row["IdTipoPlantilla"].ToString();
                            string _Fecha_Modificacion = _Row["Fecha_Modificacion"].ToString();
                            string _Linea = null;

                            if (File.Exists(_Path + "\\jsons\\Update" + _IdTipoPlantilla + ".txt"))
                            {
                                StreamReader sw = new StreamReader(_Path + "\\Jsons\\Update" + _IdTipoPlantilla + ".txt");
                                _Linea = sw.ReadLine();
                                sw.Close();

                                if (_Linea != null)
                                {
                                    if (_Linea != _Fecha_Modificacion)
                                    {
                                        this.TopMost = false;
                                        this.Enabled = false;
                                        this.Hide();
                                        FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                                        _FileJsonfrm._Form = this;
                                        _FileJsonfrm._Process = false;
                                        _FileJsonfrm._Update = true;
                                        _FileJsonfrm._window = this.Text;
                                        _FileJsonfrm.Show();
                                        return;
                                    }
                                }
                            }
                        }
                    }
                    FillTemplateType(cmbTipoPlantilla);
                    FillYears(cmbAnio);
                }
                else
                {
                    if (!_Connection)
                    {
                        MessageBox.Show(_Message.Replace("...", ", para crear los archivos base..."), "Creación de Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.btnCargar.Enabled = false;
                        return;
                    }
                    else
                    {
                        this.TopMost = false;
                        this.Enabled = false;
                        this.Hide();
                        FileJsonTemplate _FileJsonfrm = new FileJsonTemplate();
                        _FileJsonfrm._Form = this;
                        _FileJsonfrm._Process = false;
                        _FileJsonfrm._Update = false;
                        _FileJsonfrm._window = this.Text;
                        _FileJsonfrm.Show();
                        return;
                    }
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