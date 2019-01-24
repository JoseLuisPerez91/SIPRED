using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using Microsoft.Office.Core;


namespace ExcelAddIn1
{
    public partial class frmNewIndices : Form
    {
        int NroPrincipal = 0; bool ConFormula;
        public frmNewIndices(int NroFilaPrincipal, bool tieneformula)
        {
            InitializeComponent();
            txtCantIndices.Select();
            NroPrincipal = NroFilaPrincipal;
            ConFormula = tieneformula;
          //  string JsonIndice = System.IO.File.ReadAllText(@"C:\Users\AMS - 5\Indice.json");

            
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

                    ExcelAddIn1.BusinessLogic.InsertIndice(NewActiveWorksheet, cantRows, currentCell, ConFormula, NroPrincipal);
                    this.Close();
                }else
                    MessageBox.Show("Especifique por favor un dato válido.", "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else
                MessageBox.Show("Especifique por favor la cantidad de índices a insertar.", "Agregar índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);



            //NewActiveWorksheet.Application.CommandBars["Ply"].Controls["&Delete"].Enabled = false;
        }
       
        private void txtCantIndices_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
                 e.Handled = false;
            else if (Char.IsControl(e.KeyChar))
                 e.Handled = false;
            else if (Char.IsSeparator(e.KeyChar))
                 e.Handled = false;
            else
                e.Handled = true;
                    
        }
    }
}
