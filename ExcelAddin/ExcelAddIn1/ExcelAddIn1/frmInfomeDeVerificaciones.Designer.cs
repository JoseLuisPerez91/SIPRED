namespace ExcelAddIn1
{
    partial class frmInfomeDeVerificaciones
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.mnu_InformeVerifiacion = new System.Windows.Forms.MenuStrip();
            this.mItem_Detalle = new System.Windows.Forms.ToolStripMenuItem();
            this.mItem_Cerrar = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.cmb_Vista = new System.Windows.Forms.ComboBox();
            this.txt_TotalCrucesProcesados = new System.Windows.Forms.TextBox();
            this.txt_TotalCruces = new System.Windows.Forms.TextBox();
            this.dgv_Cruce = new System.Windows.Forms.DataGridView();
            this.txt_Formulas = new System.Windows.Forms.TextBox();
            this.dgv_Indice = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_TotalLadoIzq = new System.Windows.Forms.TextBox();
            this.txt_TotalLadoDer = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.mnu_VistaDeInfome = new System.Windows.Forms.MenuStrip();
            this.mItem_Cruces = new System.Windows.Forms.ToolStripMenuItem();
            this.mItem_Formulas = new System.Windows.Forms.ToolStripMenuItem();
            this.mItemValidacionS = new System.Windows.Forms.ToolStripMenuItem();
            this.mItem_ValidacionA = new System.Windows.Forms.ToolStripMenuItem();
            this.Id_Cruce = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col_Concepto = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col_Diferencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col_Nota = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Con_TieneNota = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Col_TipoMov = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col_Indice = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col_ConceptoIndice = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col_Columna = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col_Grupo1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Col_Grupo2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.mnu_InformeVerifiacion.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Cruce)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Indice)).BeginInit();
            this.panel1.SuspendLayout();
            this.mnu_VistaDeInfome.SuspendLayout();
            this.SuspendLayout();
            // 
            // mnu_InformeVerifiacion
            // 
            this.mnu_InformeVerifiacion.GripStyle = System.Windows.Forms.ToolStripGripStyle.Visible;
            this.mnu_InformeVerifiacion.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mItem_Detalle,
            this.mItem_Cerrar});
            this.mnu_InformeVerifiacion.Location = new System.Drawing.Point(0, 0);
            this.mnu_InformeVerifiacion.Name = "mnu_InformeVerifiacion";
            this.mnu_InformeVerifiacion.Size = new System.Drawing.Size(684, 24);
            this.mnu_InformeVerifiacion.TabIndex = 0;
            this.mnu_InformeVerifiacion.Text = "Menú";
            // 
            // mItem_Detalle
            // 
            this.mItem_Detalle.Name = "mItem_Detalle";
            this.mItem_Detalle.Size = new System.Drawing.Size(55, 20);
            this.mItem_Detalle.Text = "Detalle";
            this.mItem_Detalle.Click += new System.EventHandler(this.mItem_Detalle_Click);
            // 
            // mItem_Cerrar
            // 
            this.mItem_Cerrar.Name = "mItem_Cerrar";
            this.mItem_Cerrar.Size = new System.Drawing.Size(51, 20);
            this.mItem_Cerrar.Text = "Cerrar";
            this.mItem_Cerrar.Click += new System.EventHandler(this.mItem_Cerrar_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(4, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "Mostrar";
            // 
            // cmb_Vista
            // 
            this.cmb_Vista.Items.AddRange(new object[] {
            "Diferencias",
            "Correctos",
            "No Aplican"});
            this.cmb_Vista.Location = new System.Drawing.Point(63, 27);
            this.cmb_Vista.Name = "cmb_Vista";
            this.cmb_Vista.Size = new System.Drawing.Size(468, 21);
            this.cmb_Vista.TabIndex = 2;
            this.cmb_Vista.SelectedIndexChanged += new System.EventHandler(this.cmb_Vista_SelectedIndexChanged);
            // 
            // txt_TotalCrucesProcesados
            // 
            this.txt_TotalCrucesProcesados.BackColor = System.Drawing.Color.Chartreuse;
            this.txt_TotalCrucesProcesados.Location = new System.Drawing.Point(536, 28);
            this.txt_TotalCrucesProcesados.Name = "txt_TotalCrucesProcesados";
            this.txt_TotalCrucesProcesados.ReadOnly = true;
            this.txt_TotalCrucesProcesados.Size = new System.Drawing.Size(66, 20);
            this.txt_TotalCrucesProcesados.TabIndex = 3;
            // 
            // txt_TotalCruces
            // 
            this.txt_TotalCruces.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.txt_TotalCruces.Location = new System.Drawing.Point(606, 28);
            this.txt_TotalCruces.Name = "txt_TotalCruces";
            this.txt_TotalCruces.ReadOnly = true;
            this.txt_TotalCruces.Size = new System.Drawing.Size(66, 20);
            this.txt_TotalCruces.TabIndex = 4;
            // 
            // dgv_Cruce
            // 
            this.dgv_Cruce.AllowUserToAddRows = false;
            this.dgv_Cruce.AllowUserToDeleteRows = false;
            this.dgv_Cruce.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Cruce.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Id_Cruce,
            this.Col_Concepto,
            this.Col_Diferencia,
            this.Col_Nota,
            this.Con_TieneNota,
            this.Col_TipoMov});
            this.dgv_Cruce.Location = new System.Drawing.Point(5, 7);
            this.dgv_Cruce.MultiSelect = false;
            this.dgv_Cruce.Name = "dgv_Cruce";
            this.dgv_Cruce.ReadOnly = true;
            this.dgv_Cruce.RowHeadersVisible = false;
            this.dgv_Cruce.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_Cruce.Size = new System.Drawing.Size(660, 118);
            this.dgv_Cruce.TabIndex = 0;
            this.dgv_Cruce.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_Cruce_CellContentClick);
            // 
            // txt_Formulas
            // 
            this.txt_Formulas.BackColor = System.Drawing.Color.Khaki;
            this.txt_Formulas.Location = new System.Drawing.Point(5, 129);
            this.txt_Formulas.Multiline = true;
            this.txt_Formulas.Name = "txt_Formulas";
            this.txt_Formulas.ReadOnly = true;
            this.txt_Formulas.Size = new System.Drawing.Size(660, 62);
            this.txt_Formulas.TabIndex = 1;
            // 
            // dgv_Indice
            // 
            this.dgv_Indice.AllowUserToAddRows = false;
            this.dgv_Indice.AllowUserToDeleteRows = false;
            this.dgv_Indice.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_Indice.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Col_Indice,
            this.Col_ConceptoIndice,
            this.Col_Columna,
            this.Col_Grupo1,
            this.Col_Grupo2});
            this.dgv_Indice.Location = new System.Drawing.Point(5, 195);
            this.dgv_Indice.MultiSelect = false;
            this.dgv_Indice.Name = "dgv_Indice";
            this.dgv_Indice.ReadOnly = true;
            this.dgv_Indice.RowHeadersVisible = false;
            this.dgv_Indice.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_Indice.Size = new System.Drawing.Size(660, 118);
            this.dgv_Indice.TabIndex = 2;
            this.dgv_Indice.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_Indice_CellContentClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(401, 325);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 16);
            this.label2.TabIndex = 3;
            this.label2.Text = "Suma";
            // 
            // txt_TotalLadoIzq
            // 
            this.txt_TotalLadoIzq.BackColor = System.Drawing.Color.LemonChiffon;
            this.txt_TotalLadoIzq.Location = new System.Drawing.Point(450, 321);
            this.txt_TotalLadoIzq.Name = "txt_TotalLadoIzq";
            this.txt_TotalLadoIzq.Size = new System.Drawing.Size(100, 20);
            this.txt_TotalLadoIzq.TabIndex = 4;
            // 
            // txt_TotalLadoDer
            // 
            this.txt_TotalLadoDer.Location = new System.Drawing.Point(556, 321);
            this.txt_TotalLadoDer.Name = "txt_TotalLadoDer";
            this.txt_TotalLadoDer.Size = new System.Drawing.Size(100, 20);
            this.txt_TotalLadoDer.TabIndex = 5;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.panel1.Controls.Add(this.txt_TotalLadoDer);
            this.panel1.Controls.Add(this.dgv_Cruce);
            this.panel1.Controls.Add(this.txt_TotalLadoIzq);
            this.panel1.Controls.Add(this.txt_Formulas);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.dgv_Indice);
            this.panel1.Location = new System.Drawing.Point(7, 54);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(668, 347);
            this.panel1.TabIndex = 6;
            // 
            // mnu_VistaDeInfome
            // 
            this.mnu_VistaDeInfome.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.mnu_VistaDeInfome.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.mnu_VistaDeInfome.GripMargin = new System.Windows.Forms.Padding(1, 1, 0, 1);
            this.mnu_VistaDeInfome.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mItem_Cruces,
            this.mItem_Formulas,
            this.mItemValidacionS,
            this.mItem_ValidacionA});
            this.mnu_VistaDeInfome.Location = new System.Drawing.Point(0, 409);
            this.mnu_VistaDeInfome.Name = "mnu_VistaDeInfome";
            this.mnu_VistaDeInfome.Padding = new System.Windows.Forms.Padding(3, 1, 0, 1);
            this.mnu_VistaDeInfome.Size = new System.Drawing.Size(684, 24);
            this.mnu_VistaDeInfome.TabIndex = 7;
            this.mnu_VistaDeInfome.Text = "mnu_VistaDeInforme";
            // 
            // mItem_Cruces
            // 
            this.mItem_Cruces.Name = "mItem_Cruces";
            this.mItem_Cruces.Size = new System.Drawing.Size(55, 22);
            this.mItem_Cruces.Text = "Cruces";
            // 
            // mItem_Formulas
            // 
            this.mItem_Formulas.Name = "mItem_Formulas";
            this.mItem_Formulas.Size = new System.Drawing.Size(68, 22);
            this.mItem_Formulas.Text = "Fórmulas";
            // 
            // mItemValidacionS
            // 
            this.mItemValidacionS.Name = "mItemValidacionS";
            this.mItemValidacionS.Size = new System.Drawing.Size(121, 22);
            this.mItemValidacionS.Text = "Validación (SIPRED)";
            // 
            // mItem_ValidacionA
            // 
            this.mItem_ValidacionA.Name = "mItem_ValidacionA";
            this.mItem_ValidacionA.Size = new System.Drawing.Size(131, 22);
            this.mItem_ValidacionA.Text = "Validación Apéndices";
            // 
            // Id_Cruce
            // 
            this.Id_Cruce.HeaderText = "Número";
            this.Id_Cruce.Name = "Id_Cruce";
            this.Id_Cruce.ReadOnly = true;
            this.Id_Cruce.Width = 50;
            // 
            // Col_Concepto
            // 
            this.Col_Concepto.HeaderText = "Concepto";
            this.Col_Concepto.Name = "Col_Concepto";
            this.Col_Concepto.ReadOnly = true;
            this.Col_Concepto.Width = 250;
            // 
            // Col_Diferencia
            // 
            this.Col_Diferencia.HeaderText = "Diferencia";
            this.Col_Diferencia.Name = "Col_Diferencia";
            this.Col_Diferencia.ReadOnly = true;
            this.Col_Diferencia.Width = 90;
            // 
            // Col_Nota
            // 
            this.Col_Nota.HeaderText = "Nota";
            this.Col_Nota.Name = "Col_Nota";
            this.Col_Nota.ReadOnly = true;
            // 
            // Con_TieneNota
            // 
            this.Con_TieneNota.HeaderText = "Tiene Nota";
            this.Con_TieneNota.Name = "Con_TieneNota";
            this.Con_TieneNota.ReadOnly = true;
            this.Con_TieneNota.Width = 70;
            // 
            // Col_TipoMov
            // 
            this.Col_TipoMov.HeaderText = "Tipo Mov";
            this.Col_TipoMov.Name = "Col_TipoMov";
            this.Col_TipoMov.ReadOnly = true;
            this.Col_TipoMov.Width = 97;
            // 
            // Col_Indice
            // 
            this.Col_Indice.HeaderText = "Indice";
            this.Col_Indice.Name = "Col_Indice";
            this.Col_Indice.ReadOnly = true;
            // 
            // Col_ConceptoIndice
            // 
            this.Col_ConceptoIndice.HeaderText = "Concepto";
            this.Col_ConceptoIndice.Name = "Col_ConceptoIndice";
            this.Col_ConceptoIndice.ReadOnly = true;
            this.Col_ConceptoIndice.Width = 317;
            // 
            // Col_Columna
            // 
            this.Col_Columna.HeaderText = "Col.";
            this.Col_Columna.Name = "Col_Columna";
            this.Col_Columna.ReadOnly = true;
            this.Col_Columna.Width = 40;
            // 
            // Col_Grupo1
            // 
            this.Col_Grupo1.HeaderText = "Gpo. 1";
            this.Col_Grupo1.Name = "Col_Grupo1";
            this.Col_Grupo1.ReadOnly = true;
            // 
            // Col_Grupo2
            // 
            this.Col_Grupo2.HeaderText = "Gpo. 2";
            this.Col_Grupo2.Name = "Col_Grupo2";
            this.Col_Grupo2.ReadOnly = true;
            // 
            // frmInfomeDeVerificaciones
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(684, 433);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.txt_TotalCruces);
            this.Controls.Add(this.txt_TotalCrucesProcesados);
            this.Controls.Add(this.cmb_Vista);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.mnu_InformeVerifiacion);
            this.Controls.Add(this.mnu_VistaDeInfome);
            this.MainMenuStrip = this.mnu_InformeVerifiacion;
            this.MaximumSize = new System.Drawing.Size(700, 472);
            this.MinimumSize = new System.Drawing.Size(700, 472);
            this.Name = "frmInfomeDeVerificaciones";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Informe de la verificación.";
            this.Load += new System.EventHandler(this.frmInfomeDeVerificaciones_Load);
            this.mnu_InformeVerifiacion.ResumeLayout(false);
            this.mnu_InformeVerifiacion.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Cruce)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_Indice)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.mnu_VistaDeInfome.ResumeLayout(false);
            this.mnu_VistaDeInfome.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip mnu_InformeVerifiacion;
        private System.Windows.Forms.ToolStripMenuItem mItem_Detalle;
        private System.Windows.Forms.ToolStripMenuItem mItem_Cerrar;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmb_Vista;
        private System.Windows.Forms.TextBox txt_TotalCrucesProcesados;
        private System.Windows.Forms.TextBox txt_TotalCruces;
        private System.Windows.Forms.TextBox txt_TotalLadoDer;
        private System.Windows.Forms.TextBox txt_TotalLadoIzq;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridView dgv_Indice;
        private System.Windows.Forms.TextBox txt_Formulas;
        private System.Windows.Forms.DataGridView dgv_Cruce;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.MenuStrip mnu_VistaDeInfome;
        private System.Windows.Forms.ToolStripMenuItem mItem_Cruces;
        private System.Windows.Forms.ToolStripMenuItem mItem_Formulas;
        private System.Windows.Forms.ToolStripMenuItem mItemValidacionS;
        private System.Windows.Forms.ToolStripMenuItem mItem_ValidacionA;
        private System.Windows.Forms.DataGridViewTextBoxColumn Id_Cruce;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Concepto;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Diferencia;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Nota;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Con_TieneNota;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_TipoMov;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Indice;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_ConceptoIndice;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Columna;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Grupo1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Col_Grupo2;
    }
}