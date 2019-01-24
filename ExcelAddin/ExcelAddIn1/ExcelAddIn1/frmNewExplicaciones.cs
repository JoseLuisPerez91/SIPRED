﻿using System;
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
    public partial class frmNewExplicaciones : Form
    {
        public frmNewExplicaciones()
        {
            InitializeComponent();
        }

        private void btnAccept_Click(object sender, EventArgs e)
        {
            string Mensaje = string.Empty;
            Excel.Range currentCell = (Excel.Range)Globals.ThisAddIn.Application.ActiveCell.Cells;
            Excel.Worksheet NewActiveWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            //  if (currentCell.Cells[currentCell.Row+1,1]== " EXPLICACION ")
            if (TxtExplicacion.Text.Trim() == string.Empty)
                MessageBox.Show("Especifique por favor la explicación.", "Explicación índice", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
           if (TxtExplicacion.Text.Length < 100)
            {
                Mensaje = "La explicación especificada tiene " + lblcontador.Text + " caracteres, debe contener al menos 100. ¿Desea continuar ? ";

                DialogResult dialogo = MessageBox.Show(Mensaje,
                  "Explicación índice", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogo == DialogResult.Yes)
                {
                    ExcelAddIn1.BusinessLogic.InsertaExplicacion(NewActiveWorksheet, currentCell, TxtExplicacion.Text);
                    this.Close();
                }

            }
            else
                 if (TxtExplicacion.Text.Length >= 100)
                 {
                        ExcelAddIn1.BusinessLogic.InsertaExplicacion(NewActiveWorksheet, currentCell, TxtExplicacion.Text);
                        this.Close();
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