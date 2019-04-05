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
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn1
{
    public partial class ControlImprimir : UserControl
    {
        public ControlImprimir()
        {
            InitializeComponent();
        }

        lImprimir impr = new lImprimir();

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {
            impr._Imprimir(dataGridView1, checkBox3.Checked, Type.Missing);
            
        }
        private void ControlImprimir_Load(object sender, EventArgs e)
        {
            impr = new lImprimir();
            impr._CargarGrilla(dataGridView1);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            int numf = (impr._HojasSPR.Length) / impr._HojasSPR.GetLength(1);
            for (int k = 0; k < numf - 7; k++)
            {
                dataGridView1.Rows[k].Cells["Imprimir"].Value = checkBox1.Checked;
            }
        }
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            impr._PrepararImpresion(checkBox2.Checked,dataGridView1,false);
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
        }

        private void toolStripStatusLabel2_Click(object sender, EventArgs e)
        {
            printDialog1.ShowDialog();
            impr._Imprimir(dataGridView1, checkBox3.Checked, printDialog1.PrinterSettings.PrinterName);
        }

        private void toolStripStatusLabel3_Click(object sender, EventArgs e)
        {
            impr._PrepararImpresion(checkBox2.Checked, dataGridView1, true);
            var addIn = Globals.ThisAddIn;
                addIn.CerrarImprimir();

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            //Para actualizar el checkbox de la grilla
            if (dataGridView1.IsCurrentCellDirty) {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void toolStripStatusLabel4_Click(object sender, EventArgs e)
        {
            impr._Imprimir(dataGridView1, checkBox3.Checked, "PDF");

        }
    }
}
