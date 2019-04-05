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

namespace ExcelAddIn1
{
    public partial class CrucesAdmin : Base
    {
        oCruce[] _Cruces;
      
        public CrucesAdmin()
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
                if (File.Exists(_Path+ "\\jsons\\Cruces.json"))
                {                    
                     _Cruces = Assembler.LoadJson<oCruce[]>($"{_Path}\\jsons\\Cruces.json");
                     var _CrucesI = (from x in _Cruces.ToList()
                                             select new
                                                     {
                                                        Numero = x.IdCruce,
                                                         x.Concepto,
                                                         TipoMov =""
                                                     }).ToList();

                    DtCruces.DataSource=ToDataTable(_CrucesI);
                    if (DtCruces.RowCount == 0)
                        CrucesCLick(-1);
                    else
                        CrucesCLick(0);
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
                _FileJsonfrm._Update = false;
                _FileJsonfrm._window = this.Text;
                _FileJsonfrm.Show();
                return;
            }

        }
         /// <summary>
        /// Carga el Grid con los cruces del json realizando el filtro por tipo y año validando contra plantillas.json
       /// </summary>
        private void CargaDTCruces()
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
                var _CrucesI = (from x in _Cruces.ToList()
                                where x.IdTipoPlantilla == _IdTemplateType
                                select new
                                {
                                    Numero = x.IdCruce,
                                    x.Concepto,
                                    TipoMov = ""
                                }).ToList();

                if (txtbuscar.Text.Trim() != "")
                {
                    var _Crucesx = _Cruces.Where(p => p.Concepto.ToUpper().Contains(txtbuscar.Text)).ToList();
                
                     _CrucesI = (from x in _Crucesx.ToList()
                                where x.IdTipoPlantilla == _IdTemplateType
                                select new
                                {
                                    Numero = x.IdCruce,
                                    x.Concepto,
                                    TipoMov = ""
                                }).ToList();
                }

                DtCruces.DataSource = ToDataTable(_CrucesI);
            }
            else
                DtCruces.DataSource = null;

            if (DtCruces.RowCount == 0)
                CrucesCLick(-1);
            else
                CrucesCLick(0);


        }
        private void cmbTipo_SelectionChangeCommitted(object sender, EventArgs e)
        {
            CargaDTCruces();
        }

        private void btnCerrar_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
        /// <summary>
        /// Carga el Grid 2-Formulas y la descripción de la misma segun el click que haga el usuario en el Grid  1- Cruces
        /// </summary>
        /// <param name="Row"></param>
        private void CrucesCLick(int Row)
        {
            try
            {
                if (Row >= 0)
                {
                    int IDCruce = Convert.ToInt32(DtCruces.Rows[Row].Cells[0].Value.ToString());
                    int IdTp = Convert.ToInt32(cmbTipo.SelectedValue);

                    Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
                    Worksheet xlSht = null;
                    Range currentCell = null;
                    Range currentFind = null;
                    oCruce _Cruce = _Cruces.Where(o => o.IdCruce == IDCruce).FirstOrDefault();

                    _Cruce.setCeldas();
                    List<oCelda> CeldaNws = new List<oCelda>();
                    foreach (oCelda _Celda in _Cruce.CeldasFormula)
                    {
                        xlSht = (Worksheet)wb.Worksheets.get_Item(_Celda.Anexo);
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
                    txtDetalle.Text = "Cruce: " + _Cruce.Formula + System.Environment.NewLine + "Condición: " + _Cruce.Condicion;
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

        private void DtCruces_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            CrucesCLick(e.RowIndex);
        }

        private void cmbAnio_SelectionChangeCommitted(object sender, EventArgs e)
        {
            CargaDTCruces();
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
                    var MaxNro = DtCruces.Rows.Cast<DataGridViewRow>()
                                .Max(r => int.TryParse(r.Cells["Numero"].Value.ToString(), out temp) ? temp : 0);

                    ActualizarCruce form = new ActualizarCruce(temp + 1,null, cmbTipo.SelectedIndex,"A");
                    form.Text = "Agregar cruce";
                    form.ShowDialog();
                    CargaDTCruces();
          
        }

        private void btnModificar_Click(object sender, EventArgs e)
        {            
                string _Message = ((cmbTipo.SelectedIndex == 0) ? "- Debe seleccionar un tipo de plantilla" : "");
                if (_Message.Length > 0)
                {
                    MessageBox.Show(_Message, "Información Faltante", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                int Row = DtCruces.CurrentCell.RowIndex;
                int IDCruce = Convert.ToInt32(DtCruces.Rows[Row].Cells[0].Value.ToString());
                oCruce _Cruce = _Cruces.Where(o => o.IdCruce == IDCruce).FirstOrDefault();
                ActualizarCruce form = new ActualizarCruce(IDCruce, _Cruce, cmbTipo.SelectedIndex,"M");
                form.Text = "Modificar cruce";
                form.ShowDialog();
                CargaDTCruces();           
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            bool _Connection = new lSerializados().CheckConnection(Configuration.UrlConnection);
            string _Message = "No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.";
            string _Title = "Conexión de Red";
            if (_Connection)
            {
                int Row = DtCruces.CurrentCell.RowIndex;
                int IDCruce = Convert.ToInt32(DtCruces.Rows[Row].Cells[0].Value.ToString());
                DialogResult dialogo = MessageBox.Show("Desea eliminar el cruce número "+ IDCruce.ToString()+"?",
                  "Confirme", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dialogo == DialogResult.Yes)
                {
                     _Message = ((cmbTipo.SelectedIndex == 0) ? "- Debe seleccionar un tipo de plantilla" : "");
                    if (_Message.Length > 0)
                    {
                        MessageBox.Show(_Message, "Información Faltante", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    DialogResult _response = DialogResult.None;

                    oCruce _Template = new oCruce()
                    {
                        IdCruce = IDCruce,
                        IdTipoPlantilla = cmbTipo.SelectedIndex,
                        Concepto = "",
                        Formula = "",
                        Condicion = "",
                        Nota = "",
                        LecturaImportes = 0

                    };
                    KeyValuePair<bool, string[]> _result = new lCrucesAdmin(_Template,"E").Delete();
                    string _Messages = "";
                    foreach (string _Msg in _result.Value) _Messages += ((_Messages.Length > 0) ? "\r\n" : "") + _Msg;
                    if (_result.Key && _response != DialogResult.Yes) _Messages = "Cruce eliminado con éxito";
                    MessageBox.Show(_Messages, (_result.Key) ? "Proceso Existoso" : "Información Faltante", MessageBoxButtons.OK, (_result.Key) ? MessageBoxIcon.Information : MessageBoxIcon.Exclamation);
                    if (_result.Key) this.Hide();
                }
            }
            else
            {
                MessageBox.Show(_Message, _Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void txtbuscar_TextChanged(object sender, EventArgs e)
        {

            var _Crucesx = (from x in _Cruces.ToList()
                            select new
                            {
                                Numero = x.IdCruce,
                                x.Concepto,
                                TipoMov = ""
                            }).ToList();
            if (txtbuscar.Text.Trim() != "")
            {
                var _CrucesI = _Cruces.Where(p => p.Concepto.ToUpper().Contains(txtbuscar.Text)).ToList();
                 _Crucesx = (from x in _CrucesI.ToList()
                                select new
                                {
                                    Numero = x.IdCruce,
                                    x.Concepto,
                                    TipoMov = ""
                                }).ToList();
               
            }
            DtCruces.DataSource = ToDataTable(_Crucesx);
        }
    }
}
