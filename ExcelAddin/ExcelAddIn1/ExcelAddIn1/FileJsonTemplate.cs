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
using ExcelAddIn.Access;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using System.Net;

namespace ExcelAddIn1
{
    public partial class FileJsonTemplate : Base
    {
        public string _window;
        public bool _Process;
        public Form _Form;
        public FileJsonTemplate()
        {
            InitializeComponent();
        }
        private void btnGenerar_Click(object sender, EventArgs e)
        {
            List<string> _Messages = new List<string>();

            bool _Key = true;
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);

            this.pgbFile.Visible = true;
            int progress = 0;
            progress += 10;

            if (!_Connection)
            {
                for (int x = 10; x <= 100; x++)
                {
                    pgbFile.Value = 100 - x;
                    this.gbProgress.Text = "Progreso " + this.pgbFile.Value + "%";
                    System.Threading.Thread.Sleep(1500);
                    x += 10;
                }
                MessageBox.Show("No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.", "Conexión de Red", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            else
            {
                pgbFile.Value = progress;
                Invoke(new System.Action(() => this.gbProgress.Text = "Progreso " + this.pgbFile.Value + "% [Obtener Tipos de Plantillas]"));
                //System.Threading.Thread.Sleep(100);
                KeyValuePair<bool, string[]> _TiposPlantillas = new lSerializados().ObtenerTiposPlantillas();
                progress += 10;
                pgbFile.Value = progress;
                Invoke(new System.Action(() => this.gbProgress.Text = "Progreso " + this.pgbFile.Value + "% [Obtener Cruces]"));
                //System.Threading.Thread.Sleep(1000);
                KeyValuePair<bool, string[]> _Cruces = new lSerializados().ObtenerCruces();
                progress += 20;
                pgbFile.Value = progress;
                Invoke(new System.Action(() => this.gbProgress.Text = "Progreso " + this.pgbFile.Value + "% [Obtener Plantillas]"));
                //System.Threading.Thread.Sleep(1000);
                KeyValuePair<bool, string[]> _Plantillas = new lSerializados().ObtenerPlantillas();
                progress += 10;
                pgbFile.Value = progress;
                Invoke(new System.Action(() => this.gbProgress.Text = "Progreso " + this.pgbFile.Value + "% [Obtener Comprobaciones]"));
                //System.Threading.Thread.Sleep(1000);
                KeyValuePair<bool, string[]> _Comprobaciones = new lSerializados().ObtenerComprobaciones();
                progress += 20;
                pgbFile.Value = progress;
                Invoke(new System.Action(() => this.gbProgress.Text = "Progreso " + this.pgbFile.Value + "% [Obtener Validación Cruces]"));
                //System.Threading.Thread.Sleep(1000);
                KeyValuePair<bool, string[]> _Validaciones = new lSerializados().ObtenerValidacionCruces();
                progress += 10;
                pgbFile.Value = progress;
                Invoke(new System.Action(() => this.gbProgress.Text = "Progreso " + this.pgbFile.Value + "% [Obtener Indices]"));
                //System.Threading.Thread.Sleep(1000);
                KeyValuePair<bool, string[]> _Indices = new lSerializados().ObtenerIndices();
                progress += 10;
                pgbFile.Value = progress;
                Invoke(new System.Action(() => this.gbProgress.Text = "Progreso " + this.pgbFile.Value + "%"));
                //System.Threading.Thread.Sleep(1000);

                _Key = (!_TiposPlantillas.Key || !_Cruces.Key || !_Plantillas.Key || !_Comprobaciones.Key || !_Validaciones.Key || !_Indices.Key);
                _Messages.AddRange(_TiposPlantillas.Value);
                _Messages.AddRange(_Cruces.Value);
                _Messages.AddRange(_Plantillas.Value);
                _Messages.AddRange(_Comprobaciones.Value);
                _Messages.AddRange(_Validaciones.Value);
                _Messages.AddRange(_Indices.Value);
                progress += 10;
                pgbFile.Value = progress;
                Invoke(new System.Action(() => this.gbProgress.Text = "Progreso " + this.pgbFile.Value + "% [Proceso Finalizado]"));
                //System.Threading.Thread.Sleep(1000);

                string _Message = "Los Archivos fueron creados con éxito. Vuelva a cargar la pantalla de [" + _window + "]. ";

                if (_Process)
                {
                    _Message = "Los Archivos fueron creados con éxito. Click en el botón de Ok para continuar con el proceso.";
                }

                MessageBox.Show(_Message, "Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (!_Process)
                {
                    _Form.Close();
                }
                this.Close();
            }
        }
    }
}
