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
using ExcelAddIn.Access;

namespace ExcelAddIn1
{
    public partial class Indices : Base
    {
        int NroPrincipal = 0; bool ConFormula;
        public Indices(int NroFilaPrincipal, bool tieneformula)
        {
            InitializeComponent();
            txtCantIndices.Select();
            NroPrincipal = NroFilaPrincipal;
            ConFormula = tieneformula;
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            int cantRows = 0;
            
            if (txtCantIndices.Text.Trim() != string.Empty)
            {
                cantRows = Convert.ToInt32(txtCantIndices.Text);
                Excel.Worksheet NewActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

                if ((cantRows > 0) && (cantRows <= NewActiveWorksheet.Rows.Count))
                {

                    Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell.Cells;

                    NewActiveWorksheet.Unprotect(ExcelAddIn.Access.Configuration.PwsExcel);

                    Generales.InsertIndice(NewActiveWorksheet, cantRows, currentCell, ConFormula, NroPrincipal);
                    this.Close();
                }
                else
                    MessageBox.Show("Especifique por favor un dato válido.", "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                MessageBox.Show("Especifique por favor la cantidad de índices a insertar.", "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);


        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
