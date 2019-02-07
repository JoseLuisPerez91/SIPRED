using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using ExcelAddIn.Logic;
using System.Windows.Forms;

namespace ExcelAddIn1 {
    public partial class Ribbon2 {
        bool cConexion = new lSerializados().CheckConnection("http://www.google.com.mx");
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e) {

        }

        private void btnNew_Click(object sender, RibbonControlEventArgs e) {
            if (cConexion) {
                Nuevo _New = new Nuevo();
                _New.Show();
            }
            else
            {
                MessageBox.Show("No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.", "Conexión de Red", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCruces_Click(object sender, RibbonControlEventArgs e)
        {
            if (cConexion)
            {
                Cruce _Cruce = new Cruce();
                _Cruce.Show();
            }
            else
            {
                MessageBox.Show("No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.", "Conexión de Red", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnPlantilla_Click(object sender, RibbonControlEventArgs e)
        {
            if (cConexion)
            {
                LoadTemplates _Template = new LoadTemplates();
                _Template.Show();
            }
            else
            {
                MessageBox.Show("No existe conexión con el servidor de datos... Contacte a un Administrador de Red para ver las opciones de conexión.", "Conexión de Red", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}