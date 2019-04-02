namespace ExcelAddIn1
{
    partial class VerificacionDeCruce
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btn_VerificarCruceSeleccionado = new System.Windows.Forms.Button();
            this.txt_SumTotalLadoDerecho = new System.Windows.Forms.TextBox();
            this.btn_VolverAverificarCruces = new System.Windows.Forms.Button();
            this.txt_CrucesConDiferencia = new System.Windows.Forms.TextBox();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.txt_TotalCruces = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.txt_SumTotalLadoIzquierdo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dgv_LadoDerechoDeFormula = new System.Windows.Forms.DataGridView();
            this.IndiceGpo2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ConceptoGpo2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnaGpo2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DatoGpo2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dgv_LadoIzquierdoDeFormula = new System.Windows.Forms.DataGridView();
            this.Indice = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ConceptoGpo1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnaGpo1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.DatoGpo1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txt_Formula = new System.Windows.Forms.TextBox();
            this.dgv_DiferenciasEnCruces = new System.Windows.Forms.DataGridView();
            this.Numero = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Tooltip = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Concepto = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Diferencia = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lst_Anexos = new System.Windows.Forms.ListBox();
            this.btn_Informe = new System.Windows.Forms.Button();
            this.tabControl1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_LadoDerechoDeFormula)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_LadoIzquierdoDeFormula)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_DiferenciasEnCruces)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(3, 3);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(493, 563);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btn_Informe);
            this.tabPage2.Controls.Add(this.btn_VerificarCruceSeleccionado);
            this.tabPage2.Controls.Add(this.txt_SumTotalLadoDerecho);
            this.tabPage2.Controls.Add(this.btn_VolverAverificarCruces);
            this.tabPage2.Controls.Add(this.txt_CrucesConDiferencia);
            this.tabPage2.Controls.Add(this.textBox6);
            this.tabPage2.Controls.Add(this.txt_TotalCruces);
            this.tabPage2.Controls.Add(this.textBox4);
            this.tabPage2.Controls.Add(this.txt_SumTotalLadoIzquierdo);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.label2);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Controls.Add(this.txt_Formula);
            this.tabPage2.Controls.Add(this.dgv_DiferenciasEnCruces);
            this.tabPage2.Controls.Add(this.lst_Anexos);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(485, 537);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Cruces";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // btn_VerificarCruceSeleccionado
            // 
            this.btn_VerificarCruceSeleccionado.Location = new System.Drawing.Point(101, 510);
            this.btn_VerificarCruceSeleccionado.Name = "btn_VerificarCruceSeleccionado";
            this.btn_VerificarCruceSeleccionado.Size = new System.Drawing.Size(90, 22);
            this.btn_VerificarCruceSeleccionado.TabIndex = 15;
            this.btn_VerificarCruceSeleccionado.Text = "Seleccionado";
            this.btn_VerificarCruceSeleccionado.UseVisualStyleBackColor = true;
            this.btn_VerificarCruceSeleccionado.Click += new System.EventHandler(this.btn_VerificarCruceSeleccionado_Click);
            // 
            // txt_SumTotalLadoDerecho
            // 
            this.txt_SumTotalLadoDerecho.Location = new System.Drawing.Point(397, 456);
            this.txt_SumTotalLadoDerecho.Name = "txt_SumTotalLadoDerecho";
            this.txt_SumTotalLadoDerecho.Size = new System.Drawing.Size(72, 20);
            this.txt_SumTotalLadoDerecho.TabIndex = 8;
            // 
            // btn_VolverAverificarCruces
            // 
            this.btn_VolverAverificarCruces.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(0)))), ((int)(((byte)(0)))));
            this.btn_VolverAverificarCruces.ForeColor = System.Drawing.SystemColors.ButtonHighlight;
            this.btn_VolverAverificarCruces.Location = new System.Drawing.Point(3, 508);
            this.btn_VolverAverificarCruces.Name = "btn_VolverAverificarCruces";
            this.btn_VolverAverificarCruces.Size = new System.Drawing.Size(92, 26);
            this.btn_VolverAverificarCruces.TabIndex = 14;
            this.btn_VolverAverificarCruces.Text = "Verificar Todo";
            this.btn_VolverAverificarCruces.UseVisualStyleBackColor = false;
            this.btn_VolverAverificarCruces.Click += new System.EventHandler(this.btn_VolverAverificarCruces_Click);
            // 
            // txt_CrucesConDiferencia
            // 
            this.txt_CrucesConDiferencia.BackColor = System.Drawing.Color.Red;
            this.txt_CrucesConDiferencia.ForeColor = System.Drawing.Color.White;
            this.txt_CrucesConDiferencia.Location = new System.Drawing.Point(397, 482);
            this.txt_CrucesConDiferencia.Name = "txt_CrucesConDiferencia";
            this.txt_CrucesConDiferencia.Size = new System.Drawing.Size(72, 20);
            this.txt_CrucesConDiferencia.TabIndex = 13;
            // 
            // textBox6
            // 
            this.textBox6.BackColor = System.Drawing.Color.Gray;
            this.textBox6.ForeColor = System.Drawing.Color.White;
            this.textBox6.Location = new System.Drawing.Point(301, 484);
            this.textBox6.Name = "textBox6";
            this.textBox6.Size = new System.Drawing.Size(77, 20);
            this.textBox6.TabIndex = 12;
            this.textBox6.Text = "Con diferencia:";
            // 
            // txt_TotalCruces
            // 
            this.txt_TotalCruces.BackColor = System.Drawing.Color.DodgerBlue;
            this.txt_TotalCruces.ForeColor = System.Drawing.Color.White;
            this.txt_TotalCruces.Location = new System.Drawing.Point(209, 484);
            this.txt_TotalCruces.Name = "txt_TotalCruces";
            this.txt_TotalCruces.Size = new System.Drawing.Size(71, 20);
            this.txt_TotalCruces.TabIndex = 11;
            // 
            // textBox4
            // 
            this.textBox4.BackColor = System.Drawing.Color.Gray;
            this.textBox4.ForeColor = System.Drawing.Color.White;
            this.textBox4.Location = new System.Drawing.Point(101, 482);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(90, 20);
            this.textBox4.TabIndex = 10;
            this.textBox4.Text = "Evaluados:";
            // 
            // txt_SumTotalLadoIzquierdo
            // 
            this.txt_SumTotalLadoIzquierdo.Location = new System.Drawing.Point(209, 456);
            this.txt_SumTotalLadoIzquierdo.Name = "txt_SumTotalLadoIzquierdo";
            this.txt_SumTotalLadoIzquierdo.Size = new System.Drawing.Size(72, 20);
            this.txt_SumTotalLadoIzquierdo.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(333, 459);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Grupo 2";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(146, 459);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(45, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Grupo 1";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(98, 459);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Cálculos";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dgv_LadoDerechoDeFormula);
            this.groupBox2.Location = new System.Drawing.Point(101, 345);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(378, 105);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Grupo 2";
            // 
            // dgv_LadoDerechoDeFormula
            // 
            this.dgv_LadoDerechoDeFormula.AllowUserToAddRows = false;
            this.dgv_LadoDerechoDeFormula.AllowUserToDeleteRows = false;
            this.dgv_LadoDerechoDeFormula.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_LadoDerechoDeFormula.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.IndiceGpo2,
            this.ConceptoGpo2,
            this.ColumnaGpo2,
            this.DatoGpo2});
            this.dgv_LadoDerechoDeFormula.Location = new System.Drawing.Point(0, 15);
            this.dgv_LadoDerechoDeFormula.MultiSelect = false;
            this.dgv_LadoDerechoDeFormula.Name = "dgv_LadoDerechoDeFormula";
            this.dgv_LadoDerechoDeFormula.ReadOnly = true;
            this.dgv_LadoDerechoDeFormula.RowHeadersVisible = false;
            this.dgv_LadoDerechoDeFormula.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_LadoDerechoDeFormula.Size = new System.Drawing.Size(378, 84);
            this.dgv_LadoDerechoDeFormula.TabIndex = 1;
            this.dgv_LadoDerechoDeFormula.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_LadoDerechoDeFormula_CellContentClick);
            // 
            // IndiceGpo2
            // 
            this.IndiceGpo2.HeaderText = "indice";
            this.IndiceGpo2.Name = "IndiceGpo2";
            this.IndiceGpo2.ReadOnly = true;
            // 
            // ConceptoGpo2
            // 
            this.ConceptoGpo2.HeaderText = "Concepto";
            this.ConceptoGpo2.Name = "ConceptoGpo2";
            this.ConceptoGpo2.ReadOnly = true;
            this.ConceptoGpo2.Width = 150;
            // 
            // ColumnaGpo2
            // 
            this.ColumnaGpo2.HeaderText = "Col";
            this.ColumnaGpo2.Name = "ColumnaGpo2";
            this.ColumnaGpo2.ReadOnly = true;
            this.ColumnaGpo2.Width = 25;
            // 
            // DatoGpo2
            // 
            this.DatoGpo2.HeaderText = "Dato";
            this.DatoGpo2.Name = "DatoGpo2";
            this.DatoGpo2.ReadOnly = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dgv_LadoIzquierdoDeFormula);
            this.groupBox1.Location = new System.Drawing.Point(101, 227);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(378, 112);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Grupo 1";
            // 
            // dgv_LadoIzquierdoDeFormula
            // 
            this.dgv_LadoIzquierdoDeFormula.AllowUserToAddRows = false;
            this.dgv_LadoIzquierdoDeFormula.AllowUserToDeleteRows = false;
            this.dgv_LadoIzquierdoDeFormula.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_LadoIzquierdoDeFormula.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Indice,
            this.ConceptoGpo1,
            this.ColumnaGpo1,
            this.DatoGpo1});
            this.dgv_LadoIzquierdoDeFormula.Location = new System.Drawing.Point(0, 15);
            this.dgv_LadoIzquierdoDeFormula.MultiSelect = false;
            this.dgv_LadoIzquierdoDeFormula.Name = "dgv_LadoIzquierdoDeFormula";
            this.dgv_LadoIzquierdoDeFormula.ReadOnly = true;
            this.dgv_LadoIzquierdoDeFormula.RowHeadersVisible = false;
            this.dgv_LadoIzquierdoDeFormula.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_LadoIzquierdoDeFormula.Size = new System.Drawing.Size(378, 91);
            this.dgv_LadoIzquierdoDeFormula.TabIndex = 0;
            this.dgv_LadoIzquierdoDeFormula.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_LadoIzquierdoDeFormula_CellContentClick);
            // 
            // Indice
            // 
            this.Indice.HeaderText = "Indice";
            this.Indice.Name = "Indice";
            this.Indice.ReadOnly = true;
            // 
            // ConceptoGpo1
            // 
            this.ConceptoGpo1.HeaderText = "Concepto";
            this.ConceptoGpo1.Name = "ConceptoGpo1";
            this.ConceptoGpo1.ReadOnly = true;
            this.ConceptoGpo1.Width = 150;
            // 
            // ColumnaGpo1
            // 
            this.ColumnaGpo1.HeaderText = "Col";
            this.ColumnaGpo1.Name = "ColumnaGpo1";
            this.ColumnaGpo1.ReadOnly = true;
            this.ColumnaGpo1.Width = 25;
            // 
            // DatoGpo1
            // 
            this.DatoGpo1.HeaderText = "Dato";
            this.DatoGpo1.Name = "DatoGpo1";
            this.DatoGpo1.ReadOnly = true;
            // 
            // txt_Formula
            // 
            this.txt_Formula.BackColor = System.Drawing.Color.Khaki;
            this.txt_Formula.Location = new System.Drawing.Point(101, 174);
            this.txt_Formula.Multiline = true;
            this.txt_Formula.Name = "txt_Formula";
            this.txt_Formula.Size = new System.Drawing.Size(378, 47);
            this.txt_Formula.TabIndex = 2;
            // 
            // dgv_DiferenciasEnCruces
            // 
            this.dgv_DiferenciasEnCruces.AllowUserToAddRows = false;
            this.dgv_DiferenciasEnCruces.AllowUserToDeleteRows = false;
            this.dgv_DiferenciasEnCruces.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_DiferenciasEnCruces.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Numero,
            this.Tooltip,
            this.Concepto,
            this.Diferencia});
            this.dgv_DiferenciasEnCruces.Location = new System.Drawing.Point(101, 9);
            this.dgv_DiferenciasEnCruces.MultiSelect = false;
            this.dgv_DiferenciasEnCruces.Name = "dgv_DiferenciasEnCruces";
            this.dgv_DiferenciasEnCruces.ReadOnly = true;
            this.dgv_DiferenciasEnCruces.RowHeadersVisible = false;
            this.dgv_DiferenciasEnCruces.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv_DiferenciasEnCruces.Size = new System.Drawing.Size(378, 159);
            this.dgv_DiferenciasEnCruces.TabIndex = 1;
            this.dgv_DiferenciasEnCruces.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_DiferenciasEnCruces_CellContentClick);
            // 
            // Numero
            // 
            this.Numero.HeaderText = "Numero";
            this.Numero.Name = "Numero";
            this.Numero.ReadOnly = true;
            this.Numero.Width = 50;
            // 
            // Tooltip
            // 
            this.Tooltip.HeaderText = "";
            this.Tooltip.Name = "Tooltip";
            this.Tooltip.ReadOnly = true;
            this.Tooltip.Width = 30;
            // 
            // Concepto
            // 
            this.Concepto.HeaderText = "Concepto";
            this.Concepto.Name = "Concepto";
            this.Concepto.ReadOnly = true;
            this.Concepto.Width = 195;
            // 
            // Diferencia
            // 
            this.Diferencia.HeaderText = "Diferencia";
            this.Diferencia.Name = "Diferencia";
            this.Diferencia.ReadOnly = true;
            // 
            // lst_Anexos
            // 
            this.lst_Anexos.FormattingEnabled = true;
            this.lst_Anexos.Location = new System.Drawing.Point(6, 9);
            this.lst_Anexos.Name = "lst_Anexos";
            this.lst_Anexos.Size = new System.Drawing.Size(89, 498);
            this.lst_Anexos.TabIndex = 0;
            this.lst_Anexos.SelectedIndexChanged += new System.EventHandler(this.lst_Anexos_SelectedIndexChanged);
            // 
            // btn_Informe
            // 
            this.btn_Informe.Location = new System.Drawing.Point(205, 511);
            this.btn_Informe.Name = "btn_Informe";
            this.btn_Informe.Size = new System.Drawing.Size(75, 23);
            this.btn_Informe.TabIndex = 16;
            this.btn_Informe.Text = "Informe";
            this.btn_Informe.UseVisualStyleBackColor = true;
            this.btn_Informe.Click += new System.EventHandler(this.btn_Informe_Click);
            // 
            // VerificacionDeCruce
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.DimGray;
            this.Controls.Add(this.tabControl1);
            this.Name = "VerificacionDeCruce";
            this.Size = new System.Drawing.Size(499, 579);
            this.tabControl1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_LadoDerechoDeFormula)).EndInit();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_LadoIzquierdoDeFormula)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_DiferenciasEnCruces)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.TabControl tabControl1;
        public System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.TextBox txt_Formula;
        public System.Windows.Forms.DataGridView dgv_DiferenciasEnCruces;
        public System.Windows.Forms.ListBox lst_Anexos;
        public System.Windows.Forms.TextBox txt_SumTotalLadoDerecho;
        public System.Windows.Forms.Label label3;
        public System.Windows.Forms.Label label2;
        public System.Windows.Forms.Label label1;
        public System.Windows.Forms.TextBox txt_CrucesConDiferencia;
        public System.Windows.Forms.TextBox textBox6;
        public System.Windows.Forms.TextBox txt_TotalCruces;
        public System.Windows.Forms.TextBox textBox4;
        public System.Windows.Forms.TextBox txt_SumTotalLadoIzquierdo;
        public System.Windows.Forms.DataGridView dgv_LadoIzquierdoDeFormula;
        private System.Windows.Forms.GroupBox groupBox2;
        public System.Windows.Forms.DataGridView dgv_LadoDerechoDeFormula;
        private System.Windows.Forms.DataGridViewTextBoxColumn IndiceGpo2;
        private System.Windows.Forms.DataGridViewTextBoxColumn ConceptoGpo2;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnaGpo2;
        private System.Windows.Forms.DataGridViewTextBoxColumn DatoGpo2;
        private System.Windows.Forms.DataGridViewTextBoxColumn Indice;
        private System.Windows.Forms.DataGridViewTextBoxColumn ConceptoGpo1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnaGpo1;
        private System.Windows.Forms.DataGridViewTextBoxColumn DatoGpo1;
        private System.Windows.Forms.DataGridViewTextBoxColumn Numero;
        private System.Windows.Forms.DataGridViewTextBoxColumn Tooltip;
        private System.Windows.Forms.DataGridViewTextBoxColumn Concepto;
        private System.Windows.Forms.DataGridViewTextBoxColumn Diferencia;
        private System.Windows.Forms.Button btn_VolverAverificarCruces;
        private System.Windows.Forms.Button btn_VerificarCruceSeleccionado;
        private System.Windows.Forms.Button btn_Informe;
    }
}
