using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace ExcelAddIn1 {
    public partial class Ribbon2 {
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e) {

        }

        private void btnNew_Click(object sender, RibbonControlEventArgs e) {
            Nuevo _New = new Nuevo();
            _New.Show();
        }

        private void button7_Click(object sender, RibbonControlEventArgs e) {
            LoadTemplate _Template = new LoadTemplate();
            _Template.Show();
        }
    }
}