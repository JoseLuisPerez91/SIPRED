using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.IO;
using ExcelAddIn.Objects;
using ExcelAddIn.Logic;
using ExcelAddIn.Access;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace ExcelAddIn1
{
    public partial class ComprobacionesAdmin : Base
    {

        oComprobacion[] _Comprobaciones;
       
        public ComprobacionesAdmin()
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
                    FillYears(cmbAnio);
                    FillTemplateType(cmbTipo);
                }
                else
                {
                    if (!_Connection)
                    {
                        MessageBox.Show(_Message.Replace("...", ", para crear los archivos base..."), "Creación de Archivos Base", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        this.btnAgregar.Enabled = false;
                        this.btnModificar.Enabled = false;
                        this.btnEliminar.Enabled = false;
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
                if (File.Exists(_Path + "\\jsons\\Comprobaciones.json"))
                {
                    _Comprobaciones = Assembler.LoadJson<oComprobacion[]>($"{_Path}\\jsons\\Comprobaciones.json");
                    var _ComprobacionesI = (from x in _Comprobaciones.ToList()
                                           
                                    select new
                                    {
                                        Numero = x.IdComprobacion,
                                        x.Concepto,
                                        x.Nota
                                    }).ToList();

                    DtComprobaciones.DataSource = ToDataTable(_ComprobacionesI);
                    if (DtComprobaciones.RowCount == 0)
                        ComprobacionesCLick(-1);
                    else
                        ComprobacionesCLick(0);
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

        /// <summary>
        /// Carga el Grid 2-Formulas y la descripción de la misma segun el click que haga el usuario en el Grid  1- comprobaciones
        /// </summary>
        /// <param name="Row"></param>
        private void ComprobacionesCLick(int Row)
        {
            try
            {
                if (Row >= 0)
                {
                    int IDComp = Convert.ToInt32(DtComprobaciones.Rows[Row].Cells[0].Value.ToString());
                    int IdTp = Convert.ToInt32(cmbTipo.SelectedValue);

                    Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                    Range currentCell = null;
                    Range currentFind = null;
                    Worksheet xlSht = null;
                   oComprobacion _Comprobacion = _Comprobaciones.Where(o => o.IdComprobacion == IDComp).FirstOrDefault();
                    _Comprobacion.setCeldas();
                    List<oCelda> CeldaNws = new List<oCelda>();
                    foreach (oCelda _Celda in _Comprobacion.Celdas)
                    {
                         xlSht = wb.Worksheets[_Comprobacion.Destino.Anexo];
                        int _maxValue = xlSht.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

                        currentCell = (Range)xlSht.get_Range("A1", "A" + (_maxValue).ToString());
                        currentFind = currentCell.Find(_Celda.Indice, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                                                                 XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false,
                                                                  Type.Missing, Type.Missing);

                        _Celda.Fila = currentFind.Row;

                        currentCell = (Range)xlSht.Cells[_Celda.Fila, 2];
                        if (currentCell.get_Value(Type.Missing) != null)
                        {
                            _Celda.Concepto = currentCell.get_Value(Type.Missing).ToString();
                        }
                        else
                        {
                            _Celda.Concepto = "";
                        }

                        CeldaNws.Add(_Celda);
                    }

                       
                  
                    
                   

                    var _FormulaI = (from x in CeldaNws
                                     select new
                                     {
                                         Anexo = x.Anexo,
                                         Indice = x.Indice,
                                         x.Concepto,
                                         Col = Generales.ColumnAdress(x.Columna),
                                         CodSAT = ""
                                     }).ToList();
                    DtFormula.DataSource = ToDataTable(_FormulaI);
                    txtDetalle.Text = "Fórmula: " + _Comprobacion.Formula ;
                }
                else
                {
                    DtFormula.DataSource = null;
                    txtDetalle.Text = "";
                }
            }
            catch
            {
                this.Hide();
                MessageBox.Show("Archivo no válido, favor de generar el archivo mediante el AddIn D.SAT", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DtComprobaciones_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ComprobacionesCLick(e.RowIndex);
        }

        /// <summary>
        /// Carga el Grid con las fórmulas de comprobaciones del json realizando el filtro por tipo y año validando contra plantillas.json
        /// </summary>
        private void CargaDTComprobaciones()
        {
            string _Path = Configuration.Path;
            bool bOk = true;
            oPlantilla[] _Templates = Assembler.LoadJson<oPlantilla[]>($"{_Path}\\jsons\\Plantillas.json");
            int _IdTemplateType = (int)cmbTipo.SelectedValue, _Year = (int)cmbAnio.SelectedValue;

            if (_IdTemplateType == 0)
            {
                MessageBox.Show("Favor de seleccionar un Tipo de Plantilla.", "Tipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.cmbTipo.Focus();
                bOk = false;
            }
            if (_Year == 0 && bOk)
            {
                MessageBox.Show("Favor de seleccionar el Año a Aplicar al Tipo de Plantilla.", "Año", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.cmbAnio.Focus();
                bOk = false;
            }

            oPlantilla _Template = _Templates.FirstOrDefault(o => o.IdTipoPlantilla == _IdTemplateType && o.Anio == _Year);

            if (_Template == null && bOk)
            {
                MessageBox.Show("No existe una plantilla para el tipo seleccionado, favor de seleccionar otro tipo o contactar al administrador.", "Información Incorrecta", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                bOk = false;
            }
            if (bOk)
            {
                var _ComprobacionesI = (from x in _Comprobaciones.ToList()
                                where x.IdTipoPlantilla == _IdTemplateType
                                select new
                                {
                                    Numero = x.IdComprobacion,
                                    x.Concepto,
                                    x.Nota
                                }).ToList();

                if (txtbuscar.Text.Trim() != "")
                {
                    var _Comprobacionesx = _Comprobaciones.Where(p => p.Concepto.ToUpper().Contains(txtbuscar.Text)).ToList();

                    _ComprobacionesI = (from x in _Comprobacionesx.ToList()
                                where x.IdTipoPlantilla == _IdTemplateType
                                select new
                                {
                                    Numero = x.IdComprobacion,
                                    x.Concepto,
                                    x.Nota
                                }).ToList();
                }


                DtComprobaciones.DataSource = ToDataTable(_ComprobacionesI);
            }
            else
                DtComprobaciones.DataSource = null;
                

            if (DtComprobaciones.RowCount == 0)
                ComprobacionesCLick(-1);
            else
                ComprobacionesCLick(0);


        }

        private void cmbAnio_SelectionChangeCommitted(object sender, EventArgs e)
        {
            CargaDTComprobaciones();
        }

        private void txtbuscar_TextChanged(object sender, EventArgs e)
        {
            var _Comprobacionesx = (from x in _Comprobaciones.ToList()
                                     select new
                                     {
                                         Numero = x.IdComprobacion,
                                         x.Concepto,
                                         x.Nota
                                     }).ToList();

            if (txtbuscar.Text.Trim() != "")
            {
                var _ComprobacionesI = _Comprobaciones.Where(p => p.Concepto.ToUpper().Contains(txtbuscar.Text)).ToList();
                _Comprobacionesx = (from x in _ComprobacionesI.ToList()
                                select new
                                {
                                    Numero = x.IdComprobacion,
                                    x.Concepto,
                                    x.Nota
                                }).ToList();
                
            }
            DtComprobaciones.DataSource = ToDataTable(_Comprobacionesx);
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void btnAgregar_Click(object sender, EventArgs e)
        {
           
                string _Message = ((cmbTipo.SelectedIndex == 0) ? "- Debe seleccionar un tipo de plantilla" : "");
                if (_Message.Length > 0)
                {
                    MessageBox.Show(_Message, "Información Faltante", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                int temp = 0; 
                var MaxNro = DtComprobaciones.Rows.Cast<DataGridViewRow>()
                            .Max(r => int.TryParse(r.Cells["Numero"].Value.ToString(), out temp) ? temp : 0);

                ActualizarComprobacion form = new ActualizarComprobacion(temp + 1, null, cmbTipo.SelectedIndex, "A");
                form.Text = "Agregar comprobación aritmetica";
                form.ShowDialog();
                CargaDTComprobaciones();
           
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {
           
                string _Message = ((cmbTipo.SelectedIndex == 0) ? "- Debe seleccionar un tipo de plantilla" : "");

               
               if (_Message.Length > 0)
               {
                     MessageBox.Show(_Message, "Información Faltante", MessageBoxButtons.OK, MessageBoxIcon.Error);
                     return;
               }

                int Row = DtComprobaciones.CurrentCell.RowIndex;
                int IDcompro = Convert.ToInt32(DtComprobaciones.Rows[Row].Cells[0].Value.ToString());
                oComprobacion _Comprobacion = _Comprobaciones.Where(o => o.IdComprobacion == IDcompro).FirstOrDefault();

                if (_Comprobacion.AdmiteCambios == 0)
                {
                    MessageBox.Show("La fórmula no puede ser modificada ya que es un cálculo de " + cmbTipo.Text, "Modificar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
           
                    ActualizarComprobacion form = new ActualizarComprobacion(IDcompro, _Comprobacion, cmbTipo.SelectedIndex, "M");
                    form.Text = "Modificar comprobación aritmetica";
                    form.ShowDialog();
                    CargaDTComprobaciones();

          
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            string _Title = "Conexión de Red";
            if (_Connection)
            {
                int Row = DtComprobaciones.CurrentCell.RowIndex;
                int IDcompro = Convert.ToInt32(DtComprobaciones.Rows[Row].Cells[0].Value.ToString());
                oComprobacion _Comprobacion = _Comprobaciones.Where(o => o.IdComprobacion == IDcompro).FirstOrDefault();
                DialogResult dialogo = MessageBox.Show("Desea eliminar fórmula número " + IDcompro.ToString() + "?",
                  "Confirme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dialogo == DialogResult.Yes)
                {
                     _Message = ((cmbTipo.SelectedIndex == 0) ? "- Debe seleccionar un tipo de plantilla" : "");
                    if (_Message.Length > 0)
                    {
                        MessageBox.Show(_Message, "Información Faltante", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (_Comprobacion.AdmiteCambios == 0)
                    {
                        MessageBox.Show("La fórmula no puede ser eliminada ya que es un cálculo de " + cmbTipo.Text, "Eliminar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                DialogResult _response = DialogResult.None;

                    oComprobacion _Template = new oComprobacion()
                    {
                        IdComprobacion = IDcompro,
                        IdTipoPlantilla = 0,
                        Concepto = "",
                        Formula = "",
                        Condicion = "",
                        Nota = ""
                       
                    };
                    KeyValuePair<bool, string[]> _result = new lComprobacionesAdmin(_Template,"E").Delete();
                    string _Messages = "";
                    foreach (string _Msg in _result.Value) _Messages += ((_Messages.Length > 0) ? "\r\n" : "") + _Msg;
                    if (_result.Key && _response != DialogResult.Yes) _Messages = "Fórmula eliminada con éxito";
                    MessageBox.Show(_Messages, (_result.Key) ? "Proceso Existoso" : "Información Faltante", MessageBoxButtons.OK, (_result.Key) ? MessageBoxIcon.Information : MessageBoxIcon.Exclamation);
                    if (_result.Key) this.Hide();
                }
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cmbTipo_SelectionChangeCommitted(object sender, EventArgs e)
        {
            CargaDTComprobaciones();
        }

        private void DtComprobaciones_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            
            int IDComp = Convert.ToInt32(DtComprobaciones.Rows[e.RowIndex].Cells[0].Value.ToString());
            var _ComprobacionesI = _Comprobaciones.Where(p => p.IdComprobacion==IDComp).FirstOrDefault();
            if (_ComprobacionesI.AdmiteCambios==0)
            {
                DtComprobaciones.Rows[e.RowIndex].DefaultCellStyle.ForeColor = Color.Red;             
            }

        }
    }
}


