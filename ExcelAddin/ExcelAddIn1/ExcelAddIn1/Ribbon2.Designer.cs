namespace ExcelAddIn1
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
            this.btnPrellenar = this.Factory.CreateRibbonButton();
            this.btnIndice = this.Factory.CreateRibbonMenu();
            this.btnAgregarIndice = this.Factory.CreateRibbonButton();
            this.btnEliminarIndice = this.Factory.CreateRibbonButton();
            this.btnExplicacion = this.Factory.CreateRibbonMenu();
            this.btnAgregarExplicacion = this.Factory.CreateRibbonButton();
            this.btnEliminaeExplicacion = this.Factory.CreateRibbonButton();
            this.btnImprimir = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnCruces = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnConvertir = this.Factory.CreateRibbonButton();
            this.btnConvertirMas = this.Factory.CreateRibbonButton();
            this.btnTransferir = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnPlantilla = this.Factory.CreateRibbonButton();
            this.btnCrucesAdmin = this.Factory.CreateRibbonButton();
            this.btnComprobacionesAdmin = this.Factory.CreateRibbonButton();
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
            this.group1.Items.Add(this.btnPrellenar);
            this.group1.Items.Add(this.btnIndice);
            this.group1.Items.Add(this.btnExplicacion);
            this.group1.Items.Add(this.btnImprimir);
            this.group1.Label = "HOJA DE TRABAJO";
            this.group1.Name = "group1";
            // 
            // btnNew
            // 
            this.btnNew.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnNew.Image = ((System.Drawing.Image)(resources.GetObject("btnNew.Image")));
            this.btnNew.Label = "Nuevo";
            this.btnNew.Name = "btnNew";
            this.btnNew.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabView");
            this.btnNew.ScreenTip = "Nueva hoja de trabajo";
            this.btnNew.ShowImage = true;
            this.btnNew.SuperTip = "Crea una hoja de trabajo para capturar la información del cliente";
            this.btnNew.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNew_Click);
            // 
            // btnPrellenar
            // 
            this.btnPrellenar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPrellenar.Image = ((System.Drawing.Image)(resources.GetObject("btnPrellenar.Image")));
            this.btnPrellenar.Label = "Prellenar";
            this.btnPrellenar.Name = "btnPrellenar";
            this.btnPrellenar.Position = this.Factory.RibbonPosition.BeforeOfficeId("TabView");
            this.btnPrellenar.ScreenTip = "Llenar hoja de trabajo";
            this.btnPrellenar.ShowImage = true;
            this.btnPrellenar.SuperTip = "Obtiene la información seleccionada del cliente.";
            this.btnPrellenar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrellenar_Click);
            // 
            // btnIndice
            // 
            this.btnIndice.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnIndice.Image = ((System.Drawing.Image)(resources.GetObject("btnIndice.Image")));
            this.btnIndice.Items.Add(this.btnAgregarIndice);
            this.btnIndice.Items.Add(this.btnEliminarIndice);
            this.btnIndice.Label = "Índice";
            this.btnIndice.Name = "btnIndice";
            this.btnIndice.ScreenTip = "Índices";
            this.btnIndice.ShowImage = true;
            this.btnIndice.SuperTip = "Agrega o elimina índices en un anexo o apartado. Opción Agregar: Inserta una fila" +
    " para un nuevo índice debajo de la celda seleccionada.";
            // 
            // btnAgregarIndice
            // 
            this.btnAgregarIndice.Image = ((System.Drawing.Image)(resources.GetObject("btnAgregarIndice.Image")));
            this.btnAgregarIndice.Label = "Agregar";
            this.btnAgregarIndice.Name = "btnAgregarIndice";
            this.btnAgregarIndice.ScreenTip = "Agregar índice";
            this.btnAgregarIndice.ShowImage = true;
            this.btnAgregarIndice.SuperTip = "Inserta una fila para un nuevo índice debajo de la celda seleccionada. ";
            this.btnAgregarIndice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAgregarIndice_Click);
            // 
            // btnEliminarIndice
            // 
            this.btnEliminarIndice.Image = ((System.Drawing.Image)(resources.GetObject("btnEliminarIndice.Image")));
            this.btnEliminarIndice.Label = "Eliminar";
            this.btnEliminarIndice.Name = "btnEliminarIndice";
            this.btnEliminarIndice.ScreenTip = "Eliminar índice";
            this.btnEliminarIndice.ShowImage = true;
            this.btnEliminarIndice.SuperTip = "Elimina toda la fila del índice seleccionado";
            this.btnEliminarIndice.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEliminarIndice_Click);
            // 
            // btnExplicacion
            // 
            this.btnExplicacion.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExplicacion.Image = ((System.Drawing.Image)(resources.GetObject("btnExplicacion.Image")));
            this.btnExplicacion.Items.Add(this.btnAgregarExplicacion);
            this.btnExplicacion.Items.Add(this.btnEliminaeExplicacion);
            this.btnExplicacion.Label = "Explicación";
            this.btnExplicacion.Name = "btnExplicacion";
            this.btnExplicacion.ScreenTip = "Explicaciones";
            this.btnExplicacion.ShowImage = true;
            this.btnExplicacion.SuperTip = "Agrega o elimina explicaciones en un anexo o apartado. ";
            // 
            // btnAgregarExplicacion
            // 
            this.btnAgregarExplicacion.Image = ((System.Drawing.Image)(resources.GetObject("btnAgregarExplicacion.Image")));
            this.btnAgregarExplicacion.Label = "Agregar";
            this.btnAgregarExplicacion.Name = "btnAgregarExplicacion";
            this.btnAgregarExplicacion.ScreenTip = "Agregar explicación";
            this.btnAgregarExplicacion.ShowImage = true;
            this.btnAgregarExplicacion.SuperTip = "Inserta una fila de explicación debajo de la celda seleccionada. ";
            this.btnAgregarExplicacion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAgregarExplicacion_Click);
            // 
            // btnEliminaeExplicacion
            // 
            this.btnEliminaeExplicacion.Image = ((System.Drawing.Image)(resources.GetObject("btnEliminaeExplicacion.Image")));
            this.btnEliminaeExplicacion.Label = "Eliminar";
            this.btnEliminaeExplicacion.Name = "btnEliminaeExplicacion";
            this.btnEliminaeExplicacion.ScreenTip = "Eliminar explicación";
            this.btnEliminaeExplicacion.ShowImage = true;
            this.btnEliminaeExplicacion.SuperTip = "Elimina toda la fila de la explicación seleccionada.";
            this.btnEliminaeExplicacion.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEliminaeExplicacion_Click);
            // 
            // btnImprimir
            // 
            this.btnImprimir.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImprimir.Image = ((System.Drawing.Image)(resources.GetObject("btnImprimir.Image")));
            this.btnImprimir.Label = "Imprimir";
            this.btnImprimir.Name = "btnImprimir";
            this.btnImprimir.ScreenTip = "Imprimir información";
            this.btnImprimir.ShowImage = true;
            this.btnImprimir.SuperTip = "Identifica los anexos que tienen información generando una vista de impresión.";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnCruces);
            this.group2.Label = "VERIFICACIONES";
            this.group2.Name = "group2";
            // 
            // btnCruces
            // 
            this.btnCruces.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCruces.Image = ((System.Drawing.Image)(resources.GetObject("btnCruces.Image")));
            this.btnCruces.Label = "Cruces";
            this.btnCruces.Name = "btnCruces";
            this.btnCruces.ScreenTip = "Verificar";
            this.btnCruces.ShowImage = true;
            this.btnCruces.SuperTip = "Realiza la verificación de cruces entre apartados o anexos.";
            this.btnCruces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCruces_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnConvertir);
            this.group3.Items.Add(this.btnConvertirMas);
            this.group3.Items.Add(this.btnTransferir);
            this.group3.Label = "HERRAMIENTAS SAT";
            this.group3.Name = "group3";
            // 
            // btnConvertir
            // 
            this.btnConvertir.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConvertir.Image = ((System.Drawing.Image)(resources.GetObject("btnConvertir.Image")));
            this.btnConvertir.Label = "Convertir";
            this.btnConvertir.Name = "btnConvertir";
            this.btnConvertir.ScreenTip = "Convertir información";
            this.btnConvertir.ShowImage = true;
            this.btnConvertir.SuperTip = "Convierte un dictamen del año anterior a su equivalente para el presente ejercici" +
    "o";
            this.btnConvertir.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConvertir_Click);
            // 
            // btnConvertirMas
            // 
            this.btnConvertirMas.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConvertirMas.Image = ((System.Drawing.Image)(resources.GetObject("btnConvertirMas.Image")));
            this.btnConvertirMas.Label = "Conversión Masiva";
            this.btnConvertirMas.Name = "btnConvertirMas";
            this.btnConvertirMas.ScreenTip = "Convertit información";
            this.btnConvertirMas.ShowImage = true;
            this.btnConvertirMas.SuperTip = "Convierte todos los dictámenes del año anterior ubicados en una carpeta a su equi" +
    "valente para el presente ejercicio";
            // 
            // btnTransferir
            // 
            this.btnTransferir.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTransferir.Image = ((System.Drawing.Image)(resources.GetObject("btnTransferir.Image")));
            this.btnTransferir.Label = "Transferir";
            this.btnTransferir.Name = "btnTransferir";
            this.btnTransferir.ScreenTip = "Transferir información";
            this.btnTransferir.ShowImage = true;
            this.btnTransferir.SuperTip = "Transfiere la información a la plantilla .xspr";
            this.btnTransferir.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTransferir_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnPlantilla);
            this.group4.Items.Add(this.btnCrucesAdmin);
            this.group4.Items.Add(this.btnComprobacionesAdmin);
            this.group4.Label = "ADMINISTRACIÓN";
            this.group4.Name = "group4";
            // 
            // btnPlantilla
            // 
            this.btnPlantilla.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPlantilla.Image = ((System.Drawing.Image)(resources.GetObject("btnPlantilla.Image")));
            this.btnPlantilla.Label = "Plantilla SAT";
            this.btnPlantilla.Name = "btnPlantilla";
            this.btnPlantilla.ScreenTip = "Cargar plantilla";
            this.btnPlantilla.ShowImage = true;
            this.btnPlantilla.SuperTip = "Permite cargar las plantillas formuladas SIPRED.";
            this.btnPlantilla.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPlantilla_Click);
            // 
            // btnCrucesAdmin
            // 
            this.btnCrucesAdmin.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCrucesAdmin.Image = ((System.Drawing.Image)(resources.GetObject("btnCrucesAdmin.Image")));
            this.btnCrucesAdmin.Label = "Cruces";
            this.btnCrucesAdmin.Name = "btnCrucesAdmin";
            this.btnCrucesAdmin.ScreenTip = "Administrar cruces";
            this.btnCrucesAdmin.ShowImage = true;
            this.btnCrucesAdmin.SuperTip = "Permite crear, adecuar o eliminar los cruces definidos en el sistema.";
            // 
            // btnComprobacionesAdmin
            // 
            this.btnComprobacionesAdmin.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnComprobacionesAdmin.Image = ((System.Drawing.Image)(resources.GetObject("btnComprobacionesAdmin.Image")));
            this.btnComprobacionesAdmin.Label = "Comprobaciones";
            this.btnComprobacionesAdmin.Name = "btnComprobacionesAdmin";
            this.btnComprobacionesAdmin.ScreenTip = "Administrar comprobaciones";
            this.btnComprobacionesAdmin.ShowImage = true;
            this.btnComprobacionesAdmin.SuperTip = "Permite crear, adecuar o eliminar las comprobaciones aritméticas definidas en el " +
    "sistema.";
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
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImprimir;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCruces;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertir;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertirMas;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPlantilla;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCrucesAdmin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnComprobacionesAdmin;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu btnExplicacion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAgregarExplicacion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEliminaeExplicacion;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu btnIndice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAgregarIndice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEliminarIndice;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTransferir;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrellenar;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon2 Ribbon2
        {
            get { return this.GetRibbon<Ribbon2>(); }
        }
    }
}
