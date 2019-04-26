using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class VerificacionDeCruce : UserControl
    {
        static Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
        static Worksheet activeSheet = wb.Application.ActiveSheet;
        static Microsoft.Office.Interop.Excel.Range _Range = activeSheet.get_Range("B3");
        String _ValorAnterior = String.Empty;

        public VerificacionDeCruce()
        {
            InitializeComponent();
        }

        #region EVENTOS
        private void btn_VolverAverificarCruces_Click(object sender, EventArgs e)
        {
            Ribbon2 r = new Ribbon2();
            r.btnCruces_Click(null, null);
        }

        private void btn_VerificarCruceSeleccionado_Click(object sender, EventArgs e)
        {
            try
            {
                if (dgv_DiferenciasEnCruces.Rows.Count > 0)
                {
                    int _IdCruce = Convert.ToInt16(dgv_DiferenciasEnCruces.CurrentRow.Cells[0].Value.ToString());
                    var _FormulaCruce = (from item in Globals.ThisAddIn._result
                                         where item.IdCruce == _IdCruce
                                         select item.FormulaExcel
                                         ).ToList();

                    _ValorAnterior = _Range.get_Value(Type.Missing);
                    var _SplitFormulaCruce = _FormulaCruce[0].Split('=');
                    String _Diferencia = String.Empty;

                    if (_FormulaCruce.Count() > 0)
                    {
                        _Range.NumberFormat = "0.00";
                        _Range.Formula = "=(" + _SplitFormulaCruce[0] + "-" + _SplitFormulaCruce[1] + ")";
                        _Diferencia = _Range.get_Value(Type.Missing).ToString();                      
                    }

                    if ((_Diferencia != String.Empty && _Diferencia != null))
                    {
                        MessageBox.Show($"El cruce {_IdCruce.ToString()} tiene una dierencia de: {_Diferencia}", "Verificación del cruce seleccionado", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    else
                    {
                        MessageBox.Show("Hubo un error al calcular la diferencia. Por favor intente de nuevo. [_Diferecia NULL OR Empty]");
                    }
                }
                else
                {
                    MessageBox.Show($"No hay datos a evaluar", "Sin datos", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al evaluar el cruce [VerificacionDeCruce].[btn_VerificarCruceSeleccionado].[225] : {ex.Message}");
            }
            finally
            {                           
                _Range.NumberFormat = "@";
                _Range.Value = "";                
                _Range.Value = _ValorAnterior;
                _ValorAnterior = String.Empty;
              
            }

        }

        private void dgv_LadoDerechoDeFormula_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            String _Indice, _Valor;
            int _IdCruce, _Columna;

            _IdCruce = Convert.ToInt16(dgv_DiferenciasEnCruces.CurrentRow.Cells[0].Value.ToString());
            _Indice = dgv_LadoDerechoDeFormula.CurrentRow.Cells[0].Value.ToString();
            _Valor = dgv_LadoDerechoDeFormula.CurrentRow.Cells[3].Value.ToString();
            _Columna = Convert.ToInt16(dgv_LadoDerechoDeFormula.CurrentRow.Cells[2].Value.ToString());

            SwitchHojaYCelda(_IdCruce, _Indice, _Valor, _Columna);
        }

       
        private void dgv_LadoIzquierdoDeFormula_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            String _Indice, _Valor;
            int _IdCruce, _Columna;

            _IdCruce = Convert.ToInt16(dgv_DiferenciasEnCruces.CurrentRow.Cells[0].Value.ToString());
            _Indice = dgv_LadoIzquierdoDeFormula.CurrentRow.Cells[0].Value.ToString();
            _Valor = dgv_LadoIzquierdoDeFormula.CurrentRow.Cells[3].Value.ToString();
            _Columna = Convert.ToInt16(dgv_LadoIzquierdoDeFormula.CurrentRow.Cells[2].Value.ToString());

            SwitchHojaYCelda(_IdCruce, _Indice, _Valor, _Columna);

        }

        private void dgv_DiferenciasEnCruces_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            String _Indice, _Valor; ;
            int _IdCruce = Convert.ToInt16(dgv_DiferenciasEnCruces.CurrentRow.Cells[0].Value.ToString());
            int _Columna;

            FillTablasDeIndices(_IdCruce);

            //_Anexo = lst_Anexos.SelectedItem.ToString();
            
            _Indice = dgv_LadoIzquierdoDeFormula.CurrentRow.Cells[0].Value.ToString();
            _Columna = Convert.ToInt16(dgv_LadoIzquierdoDeFormula.CurrentRow.Cells[2].Value.ToString());

            if (dgv_LadoIzquierdoDeFormula.CurrentRow.Cells[3].Value.Equals(""))
            {
                _Valor = "0";
            }
            else
            {
                _Valor = dgv_LadoIzquierdoDeFormula.CurrentRow.Cells[3].Value.ToString();
            }

            SwitchHojaYCelda( _IdCruce, _Indice, _Valor, _Columna);
            
        }
        
        private void lst_Anexos_SelectedIndexChanged(object sender, EventArgs e)
        {
            FillTablaDeDiferenciasByAnexo(lst_Anexos.SelectedItem.ToString());
            dgv_DiferenciasEnCruces_CellContentClick(null,null);


        }
        private void btn_Informe_Click(object sender, EventArgs e)
        {
            frmInfomeDeVerificaciones Informe = new frmInfomeDeVerificaciones();
            Informe.ShowDialog();
        }
        #endregion

        #region MÉTODOS Y FUNCIONES
        private void FillTablaDeDiferenciasByAnexo(String _Anexo)
        {
            try
            {
                dgv_DiferenciasEnCruces.DataSource = null;
                dgv_DiferenciasEnCruces.Rows.Clear();
                var _DiferenciasPorAnexo = (from items in Globals.ThisAddIn._result
                                            from details in items.CeldasFormula
                                            where details.Anexo == _Anexo
                                            select new
                                            {
                                                IdCruce = items.IdCruce,
                                                Concepto = items.Concepto,
                                                Diferencia = items.Diferencia
                                            }).Distinct().ToList();

                foreach (var dif in _DiferenciasPorAnexo)
                {
                    dgv_DiferenciasEnCruces.Rows.Add(dif.IdCruce.ToString(), "", dif.Concepto, dif.Diferencia.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hubo un error al cargar la tabla de conceptos con diferencia, por favor intende de nuevo. Si el problema persiste póngase en contacto con soporte:  {ex.Message}",
                                "ERROR",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        private void FillTablasDeIndices(int _IdCruce)
        {
            try
            {
                dgv_LadoIzquierdoDeFormula.DataSource = null;
                dgv_LadoIzquierdoDeFormula.Rows.Clear();

                dgv_LadoDerechoDeFormula.DataSource = null;
                dgv_LadoDerechoDeFormula.Rows.Clear();

                var _Indices = (from items in Globals.ThisAddIn._result
                                from details in items.CeldasFormula
                                where items.IdCruce == _IdCruce 
                                select new
                                {
                                    Indice = details.Indice,
                                    Original = details.Original,
                                    Concepto = details.Concepto,
                                    Columna = details.Columna,
                                    Dato = details.Valor,
                                    Grupo = details.Grupo,
                                    Grupo1 = items.Grupo1,
                                    Grupo2 = items.Grupo2,
                                    Formula = items.Formula,
                                    Condicion = items.Condicion
                                }).ToList();

                foreach (var indice in _Indices)
                {
                    String[] _Splitformula = indice.Formula.Split('=');

                    if (_Splitformula[0].Contains(indice.Original.ToString()))
                    {
                        dgv_LadoIzquierdoDeFormula.Rows.Add(indice.Indice, indice.Concepto, indice.Columna, indice.Dato);
                    }
                    if (_Splitformula[1].Contains(indice.Original.ToString()))
                    {
                        dgv_LadoDerechoDeFormula.Rows.Add(indice.Indice, indice.Concepto, indice.Columna, indice.Dato);
                    }
                }

                txt_Formula.Text = _Indices[0].Formula;
                txt_Formula.AppendText("\r\n" + _Indices[0].Condicion);
                txt_SumTotalLadoIzquierdo.Text = _Indices[0].Grupo1;
                txt_SumTotalLadoDerecho.Text = _Indices[0].Grupo2;

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hubo un error al cargar las tablas de indices, por favor intente de nuevo. Si el problema persiste póngase en contacto con soporte: {ex.Message}");
            }
        }

        private void SwitchHojaYCelda( int IdCruce, String Indice, String Valor, int Columna)
        {
            
            string[] celValue = new string[1];
            String _Anexo = String.Empty;

            var _indicesCruce = (from items in Globals.ThisAddIn._result
                                 from details in items.CeldasFormula
                                 where items.IdCruce == IdCruce
                                 select new
                                 {
                                     Indice = details.Indice,                                     
                                     Anexo = details.Anexo,
                                     Columna = details.Columna,
                                     Fila = details.Fila,
                                     Celda = details.CeldaExcel,
                                     Valor = details.Valor
                                 }).ToList();

            var _celdaExcel = (from details in _indicesCruce
                               where details.Indice == Indice && details.Valor==Valor && details.Columna==Columna
                               select new
                               {
                                   Workbook = details.Anexo,
                                   Celda = details.Celda
                               }
                               ).ToList();
                                  
            if (_celdaExcel.Count > 0)
            {
                _Anexo = _celdaExcel[0].Workbook;
                celValue = _celdaExcel[0].Celda.Split('!');

                ((Microsoft.Office.Interop.Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[_Anexo]).Select();
                Microsoft.Office.Interop.Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
                var range = activeSheet.get_Range(celValue[1]);
                range.Select();
            }         
        }
        #endregion        
    }
}
