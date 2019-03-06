using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using ExcelAddIn.Access;

namespace ExcelAddIn1
{
    public partial class ActualizarCruce : Form
    {
        string accion = "";
        int TpPlantilla = 0;
        public ActualizarCruce(int IdCruce, oCruce _Cruce, int IDPlantilla, string Accion)
        {
            InitializeComponent();
            txtNro.Text = IdCruce.ToString();
           if (_Cruce!=null)
           {
                txtConcepto.Text = _Cruce.Concepto;
                string[] Formula = _Cruce.Formula.Split('=');
                txtcruzar.Text = Formula[0];
                txtcontra.Text = Formula[1];
                txtCondicion.Text = _Cruce.Condicion;
                if (txtCondicion.Text.Trim()!="")
                {
                   txtCondicion.ReadOnly = false;
                    chkCondicionar.Checked = true;
                }
                txtNota.Text = _Cruce.Nota;
                rbConsideraSigno.Checked = (_Cruce.LecturaImportes == 1);
                rbVabsoluto.Checked = (_Cruce.LecturaImportes == 0);


            }
            accion = Accion;
            
            TpPlantilla = IDPlantilla;
        }

        private void btguardar_Click(object sender, EventArgs e)
        {
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            string _Title = "Conexión de Red";
            if (_Connection)
            {
                _Message = (txtConcepto.Text.Trim() == "") ? "- Debe indicar concepto." : "";
                _Message += (txtcruzar.Text.Trim() == "") ? ((_Message.Length > 0) ? "\r\n" : "") + "- Debe indicar cruzar." : "";
                _Message += (txtcontra.Text.Trim() == "") ? ((_Message.Length > 0) ? "\r\n" : "") + "- Debe indicar contra." : "";
                _Message += (chkCondicionar.Checked && txtCondicion.Text.Trim() == "") ? ((_Message.Length > 0) ? "\r\n" : "") + "- Debe indicar condición." : "";

                if (_Message.Length > 0)
                {
                    MessageBox.Show(_Message, "Información Faltante", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                DialogResult _response = DialogResult.None;
                string Formulax = (txtcruzar.Text + "=" + txtcontra.Text);
                string condicion = "";
                int Valor = 0;
                if (chkCondicionar.Checked)
                    condicion = txtCondicion.Text;
                if (rbConsideraSigno.Checked)
                    Valor = 1;

                oCruce _Template = new oCruce()
                {
                    IdCruce = Convert.ToInt32(txtNro.Text),
                    IdTipoPlantilla = TpPlantilla,//buscar
                    Concepto = txtConcepto.Text,
                    Formula = Formulax,
                    Condicion = condicion,
                    Nota = txtNota.Text,
                    LecturaImportes = Valor

                };

                if (accion == "A")
                {
                    KeyValuePair<bool, string[]> _result = new lCrucesAdmin(_Template, accion).Add();

                    string _Messages = "";
                    foreach (string _Msg in _result.Value) _Messages += ((_Messages.Length > 0) ? "\r\n" : "") + _Msg;
                    if (_result.Key && _response != DialogResult.Yes) _Messages = "Cruce agregado con éxito";
                    MessageBox.Show(_Messages, (_result.Key) ? "Proceso Existoso" : "Información Faltante", MessageBoxButtons.OK, (_result.Key) ? MessageBoxIcon.Information : MessageBoxIcon.Exclamation);
                    if (_result.Key) this.Hide();
                }
                else
                     if (accion == "M")
                {
                    KeyValuePair<bool, string[]> _result = new lCrucesAdmin(_Template, accion).Update();

                    string _Messages = "";
                    foreach (string _Msg in _result.Value) _Messages += ((_Messages.Length > 0) ? "\r\n" : "") + _Msg;
                    if (_result.Key && _response != DialogResult.Yes) _Messages = "Cruce modificado con éxito";
                    MessageBox.Show(_Messages, (_result.Key) ? "Proceso Existoso" : "Información Faltante", MessageBoxButtons.OK, (_result.Key) ? MessageBoxIcon.Information : MessageBoxIcon.Exclamation);
                    if (_result.Key) this.Hide();
                }
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chkCondicionar_CheckedChanged(object sender, EventArgs e)
        {
            txtCondicion.ReadOnly = !chkCondicionar.Checked;
            if (!chkCondicionar.Checked)
                txtCondicion.Text = "";
        }

        private void txtNota_TextChanged(object sender, EventArgs e)
        {
            if (txtNota.Text.Length>500)
            {
                MessageBox.Show("Debe ingresar máximo 500 caracteres para la nota", "Error" , MessageBoxButtons.OK,   MessageBoxIcon.Exclamation);
            }
        }

        private void btncancelar_Click(object sender, EventArgs e)
        {            
            this.Hide();
        }

        private void txtNota_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }
    }
}
