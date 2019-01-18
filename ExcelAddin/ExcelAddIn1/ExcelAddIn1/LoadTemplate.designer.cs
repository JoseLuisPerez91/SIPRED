namespace ExcelAddIn1 {
    partial class LoadTemplate {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if(disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            this.ofdTemplate = new System.Windows.Forms.OpenFileDialog();
            this.label1 = new System.Windows.Forms.Label();
            this.gbTipo = new System.Windows.Forms.GroupBox();
            this.cmbAnio = new System.Windows.Forms.ComboBox();
            this.cmbTipoPlantilla = new System.Windows.Forms.ComboBox();
            this.lblPlantilla = new System.Windows.Forms.Label();
            this.txtPlantilla = new System.Windows.Forms.TextBox();
            this.btnSeleccionar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnCargar = new System.Windows.Forms.Button();
            this.gbTipo.SuspendLayout();
            this.SuspendLayout();
            // 
            // ofdTemplate
            // 
            this.ofdTemplate.Filter = "SAT Template | *.xlsm";
            this.ofdTemplate.FileOk += new System.ComponentModel.CancelEventHandler(this.ofdTemplate_FileOk);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Location = new System.Drawing.Point(16, 11);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(568, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "SELECCIONE EL TIPO Y EJERCICION DEL ARCHIVO.";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // gbTipo
            // 
            this.gbTipo.Controls.Add(this.cmbAnio);
            this.gbTipo.Controls.Add(this.cmbTipoPlantilla);
            this.gbTipo.Location = new System.Drawing.Point(16, 43);
            this.gbTipo.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.gbTipo.Name = "gbTipo";
            this.gbTipo.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.gbTipo.Size = new System.Drawing.Size(568, 68);
            this.gbTipo.TabIndex = 2;
            this.gbTipo.TabStop = false;
            this.gbTipo.Text = "Tipo";
            // 
            // cmbAnio
            // 
            this.cmbAnio.FormattingEnabled = true;
            this.cmbAnio.Location = new System.Drawing.Point(468, 25);
            this.cmbAnio.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbAnio.Name = "cmbAnio";
            this.cmbAnio.Size = new System.Drawing.Size(91, 24);
            this.cmbAnio.TabIndex = 1;
            // 
            // cmbTipoPlantilla
            // 
            this.cmbTipoPlantilla.FormattingEnabled = true;
            this.cmbTipoPlantilla.Location = new System.Drawing.Point(9, 25);
            this.cmbTipoPlantilla.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbTipoPlantilla.Name = "cmbTipoPlantilla";
            this.cmbTipoPlantilla.Size = new System.Drawing.Size(449, 24);
            this.cmbTipoPlantilla.TabIndex = 0;
            // 
            // lblPlantilla
            // 
            this.lblPlantilla.AutoSize = true;
            this.lblPlantilla.Location = new System.Drawing.Point(16, 119);
            this.lblPlantilla.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblPlantilla.Name = "lblPlantilla";
            this.lblPlantilla.Size = new System.Drawing.Size(57, 17);
            this.lblPlantilla.TabIndex = 3;
            this.lblPlantilla.Text = "Plantilla";
            // 
            // txtPlantilla
            // 
            this.txtPlantilla.Location = new System.Drawing.Point(16, 140);
            this.txtPlantilla.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtPlantilla.Name = "txtPlantilla";
            this.txtPlantilla.ReadOnly = true;
            this.txtPlantilla.Size = new System.Drawing.Size(459, 22);
            this.txtPlantilla.TabIndex = 4;
            // 
            // btnSeleccionar
            // 
            this.btnSeleccionar.Location = new System.Drawing.Point(484, 138);
            this.btnSeleccionar.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnSeleccionar.Name = "btnSeleccionar";
            this.btnSeleccionar.Size = new System.Drawing.Size(100, 28);
            this.btnSeleccionar.TabIndex = 5;
            this.btnSeleccionar.Text = "Seleccionar";
            this.btnSeleccionar.UseVisualStyleBackColor = true;
            this.btnSeleccionar.Click += new System.EventHandler(this.btnSeleccionar_Click);
            // 
            // btnCancelar
            // 
            this.btnCancelar.Location = new System.Drawing.Point(484, 186);
            this.btnCancelar.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(100, 28);
            this.btnCancelar.TabIndex = 6;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
            // 
            // btnCargar
            // 
            this.btnCargar.Location = new System.Drawing.Point(376, 186);
            this.btnCargar.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCargar.Name = "btnCargar";
            this.btnCargar.Size = new System.Drawing.Size(100, 28);
            this.btnCargar.TabIndex = 7;
            this.btnCargar.Text = "Cargar";
            this.btnCargar.UseVisualStyleBackColor = true;
            this.btnCargar.Click += new System.EventHandler(this.btnCargar_Click);
            // 
            // LoadTemplate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(601, 225);
            this.Controls.Add(this.btnCargar);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.btnSeleccionar);
            this.Controls.Add(this.txtPlantilla);
            this.Controls.Add(this.lblPlantilla);
            this.Controls.Add(this.gbTipo);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "LoadTemplate";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Cargar Plantilla";
            this.gbTipo.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog ofdTemplate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox gbTipo;
        private System.Windows.Forms.ComboBox cmbAnio;
        private System.Windows.Forms.ComboBox cmbTipoPlantilla;
        private System.Windows.Forms.Label lblPlantilla;
        private System.Windows.Forms.TextBox txtPlantilla;
        private System.Windows.Forms.Button btnSeleccionar;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnCargar;
    }
}