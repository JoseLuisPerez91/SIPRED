using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Core;

namespace ExcelAddIn1
{
    public partial class Explicaciones : Form
    {
        public Explicaciones()
        {
            InitializeComponent();
        }
        private void btnAccept_Click(object sender, EventArgs e)
        {           
            string Mensaje = string.Empty;
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell.Cells;
            Excel.Worksheet NewActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

            if (TxtExplicacion.Text.Trim() == string.Empty)
            {
                MessageBox.Show("Especifique por favor la explicación.", "Explicación índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
            {
                if (TxtExplicacion.Text.Length < 100)
                {
                    Mensaje = "La explicación especificada tiene " + lblcontador.Text + " caracteres, debe contener al menos 100. ¿Desea continuar ? ";

                    DialogResult dialogo = MessageBox.Show(Mensaje,
                      "Explicación índice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogo == DialogResult.Yes)
                    {
                        NewActiveWorksheet.Unprotect(ExcelAddIn.Access.Configuration.PwsExcel);
                        Generales.InsertaExplicacion(NewActiveWorksheet, currentCell, TxtExplicacion.Text);
                        NewActiveWorksheet.Protect(ExcelAddIn.Access.Configuration.PwsExcel, true, true, false, true, true, true, true, false, false, false, false, false, false, true, false);
                        this.Close();
                    }
                }
                else
                {
                    if (TxtExplicacion.Text.Length >= 100)
                    {
                        NewActiveWorksheet.Unprotect(ExcelAddIn.Access.Configuration.PwsExcel);

                        Generales.InsertaExplicacion(NewActiveWorksheet, currentCell, TxtExplicacion.Text);

                        NewActiveWorksheet.Protect(ExcelAddIn.Access.Configuration.PwsExcel, true, true, false, true, true, true, true, false, false, false, false, false, false, true, false);
                        this.Close();
                    }
                }
            }
        }
        private void TxtExplicacion_TextChanged(object sender, EventArgs e)
        {
            lblcontador.Text = TxtExplicacion.Text.Length.ToString();
        }
        private void TxtExplicacion_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.KeyChar = Char.ToUpper(e.KeyChar);
        }
    }
}
