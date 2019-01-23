﻿namespace ExcelAddIn1
{
    partial class Ribbon2 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon2()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon2));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnNew = this.Factory.CreateRibbonButton();
            this.btnIndice = this.Factory.CreateRibbonMenu();
            this.btnNewIndice = this.Factory.CreateRibbonButton();
            this.btnDelIndice = this.Factory.CreateRibbonButton();
            this.menu2 = this.Factory.CreateRibbonMenu();
            this.btnNewExpl = this.Factory.CreateRibbonButton();
            this.btnDelExpl = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "D.SAT";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnNew);
            this.group1.Items.Add(this.btnIndice);
            this.group1.Items.Add(this.menu2);
            this.group1.Items.Add(this.button2);
            this.group1.Label = "HOJA DE TRABAJO";
            this.group1.Name = "group1";
            // 
            // btnNew
            // 
            this.btnNew.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNew.Image = ((System.Drawing.Image)(resources.GetObject("btnNew.Image")));
            this.btnNew.Label = "Nuevo";
            this.btnNew.Name = "btnNew";
            this.btnNew.OfficeImageId = "HappyFace";
            this.btnNew.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabView");
            this.btnNew.ShowImage = true;
            this.btnNew.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNew_Click);
            // 
            // btnIndice
            // 
            this.btnIndice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnIndice.Image = ((System.Drawing.Image)(resources.GetObject("btnIndice.Image")));
            this.btnIndice.Items.Add(this.btnNewIndice);
            this.btnIndice.Items.Add(this.btnDelIndice);
            this.btnIndice.Label = "Índice";
            this.btnIndice.Name = "btnIndice";
            this.btnIndice.ShowImage = true;
            // 
            // btnNewIndice
            // 
            this.btnNewIndice.Image = ((System.Drawing.Image)(resources.GetObject("btnNewIndice.Image")));
            this.btnNewIndice.Label = "Agregar";
            this.btnNewIndice.Name = "btnNewIndice";
            this.btnNewIndice.ShowImage = true;
            this.btnNewIndice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewIndice_Click);
            // 
            // btnDelIndice
            // 
            this.btnDelIndice.Image = ((System.Drawing.Image)(resources.GetObject("btnDelIndice.Image")));
            this.btnDelIndice.Label = "Eliminar";
            this.btnDelIndice.Name = "btnDelIndice";
            this.btnDelIndice.ShowImage = true;
            this.btnDelIndice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDelIndice_Click);
            // 
            // menu2
            // 
            this.menu2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu2.Image = ((System.Drawing.Image)(resources.GetObject("menu2.Image")));
            this.menu2.Items.Add(this.btnNewExpl);
            this.menu2.Items.Add(this.btnDelExpl);
            this.menu2.Label = "Explicación";
            this.menu2.Name = "menu2";
            this.menu2.ShowImage = true;
            // 
            // btnNewExpl
            // 
            this.btnNewExpl.Image = ((System.Drawing.Image)(resources.GetObject("btnNewExpl.Image")));
            this.btnNewExpl.Label = "Agregar";
            this.btnNewExpl.Name = "btnNewExpl";
            this.btnNewExpl.ShowImage = true;
            this.btnNewExpl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNewExpl_Click);
            // 
            // btnDelExpl
            // 
            this.btnDelExpl.Image = ((System.Drawing.Image)(resources.GetObject("btnDelExpl.Image")));
            this.btnDelExpl.Label = "Eliminar";
            this.btnDelExpl.Name = "btnDelExpl";
            this.btnDelExpl.ShowImage = true;
            this.btnDelExpl.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDelExpl_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Imprimir";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            // 
            // group2
            // 
            this.group2.Items.Add(this.button3);
            this.group2.Items.Add(this.button4);
            this.group2.Label = "VERIFICACIONES";
            this.group2.Name = "group2";
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Cruces";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Label = "Comprobaciones";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            // 
            // group3
            // 
            this.group3.Items.Add(this.button6);
            this.group3.Items.Add(this.button5);
            this.group3.Label = "HERRAMIENTAS SAT";
            this.group3.Name = "group3";
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.Label = "Conversión Masiva";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.Label = "Convertir";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            // 
            // group4
            // 
            this.group4.Items.Add(this.button7);
            this.group4.Items.Add(this.button8);
            this.group4.Items.Add(this.button9);
            this.group4.Label = "ADMINISTRACIÓN";
            this.group4.Name = "group4";
            // 
            // button7
            // 
            this.button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button7.Image = ((System.Drawing.Image)(resources.GetObject("button7.Image")));
            this.button7.Label = "Plantilla SAT";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            // 
            // button8
            // 
            this.button8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button8.Image = ((System.Drawing.Image)(resources.GetObject("button8.Image")));
            this.button8.Label = "Cruces";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            // 
            // button9
            // 
            this.button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button9.Image = ((System.Drawing.Image)(resources.GetObject("button9.Image")));
            this.button9.Label = "Comprobaciones";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            // 
            // Ribbon2
            // 
            this.Name = "Ribbon2";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon2_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNew;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewExpl;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDelExpl;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu btnIndice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNewIndice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDelIndice;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon2 Ribbon2
        {
            get { return this.GetRibbon<Ribbon2>(); }
        }
    }
}
