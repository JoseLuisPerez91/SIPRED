namespace ExcelAddIn1
{
    partial class LoadTemplates
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
            this.ofdTemplate = new System.Windows.Forms.OpenFileDialog();
            this.gbTipo = new System.Windows.Forms.GroupBox();
            this.cmbAnio = new System.Windows.Forms.ComboBox();
            this.cmbTipoPlantilla = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCargar = new System.Windows.Forms.Button();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnSeleccionar = new System.Windows.Forms.Button();
            this.txtPlantilla = new System.Windows.Forms.TextBox();
            this.lblPlantilla = new System.Windows.Forms.Label();
            this.gbTipo.SuspendLayout();
            this.SuspendLayout();
            // 
            // ofdTemplate
            // 
            this.ofdTemplate.Filter = "SAT Template | *.xlsm";
            // 
            // gbTipo
            // 
            this.gbTipo.Controls.Add(this.cmbAnio);
            this.gbTipo.Controls.Add(this.cmbTipoPlantilla);
            this.gbTipo.Location = new System.Drawing.Point(12, 35);
            this.gbTipo.Name = "gbTipo";
            this.gbTipo.Size = new System.Drawing.Size(426, 55);
            this.gbTipo.TabIndex = 9;
            this.gbTipo.TabStop = false;
            this.gbTipo.Text = "Tipo";
            // 
            // cmbAnio
            // 
            this.cmbAnio.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbAnio.FormattingEnabled = true;
            this.cmbAnio.Location = new System.Drawing.Point(351, 20);
            this.cmbAnio.Name = "cmbAnio";
            this.cmbAnio.Size = new System.Drawing.Size(69, 21);
            this.cmbAnio.TabIndex = 1;
            // 
            // cmbTipoPlantilla
            // 
            this.cmbTipoPlantilla.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTipoPlantilla.FormattingEnabled = true;
            this.cmbTipoPlantilla.Location = new System.Drawing.Point(7, 20);
            this.cmbTipoPlantilla.Name = "cmbTipoPlantilla";
            this.cmbTipoPlantilla.Size = new System.Drawing.Size(338, 21);
            this.cmbTipoPlantilla.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(426, 23);
            this.label1.TabIndex = 8;
            this.label1.Text = "SELECCIONE EL TIPO Y EJERCICION DEL ARCHIVO.";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnCargar
            // 
            this.btnCargar.Location = new System.Drawing.Point(282, 151);
            this.btnCargar.Name = "btnCargar";
            this.btnCargar.Size = new System.Drawing.Size(75, 23);
            this.btnCargar.TabIndex = 14;
            this.btnCargar.Text = "Cargar";
            this.btnCargar.UseVisualStyleBackColor = true;
            this.btnCargar.Click += new System.EventHandler(this.btnCargar_Click_1);
            // 
            // btnCancelar
            // 
            this.btnCancelar.Location = new System.Drawing.Point(363, 151);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(75, 23);
            this.btnCancelar.TabIndex = 13;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click_1);
            // 
            // btnSeleccionar
            // 
            this.btnSeleccionar.Location = new System.Drawing.Point(363, 112);
            this.btnSeleccionar.Name = "btnSeleccionar";
            this.btnSeleccionar.Size = new System.Drawing.Size(75, 23);
            this.btnSeleccionar.TabIndex = 12;
            this.btnSeleccionar.Text = "Seleccionar";
            this.btnSeleccionar.UseVisualStyleBackColor = true;
            this.btnSeleccionar.Click += new System.EventHandler(this.btnSeleccionar_Click_1);
            // 
            // txtPlantilla
            // 
            this.txtPlantilla.Location = new System.Drawing.Point(12, 114);
            this.txtPlantilla.Name = "txtPlantilla";
            this.txtPlantilla.ReadOnly = true;
            this.txtPlantilla.Size = new System.Drawing.Size(345, 20);
            this.txtPlantilla.TabIndex = 11;
            // 
            // lblPlantilla
            // 
            this.lblPlantilla.AutoSize = true;
            this.lblPlantilla.Location = new System.Drawing.Point(12, 97);
            this.lblPlantilla.Name = "lblPlantilla";
            this.lblPlantilla.Size = new System.Drawing.Size(43, 13);
            this.lblPlantilla.TabIndex = 10;
            this.lblPlantilla.Text = "Plantilla";
            // 
            // LoadTemplates
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(451, 183);
            this.Controls.Add(this.gbTipo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnCargar);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.btnSeleccionar);
            this.Controls.Add(this.txtPlantilla);
            this.Controls.Add(this.lblPlantilla);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LoadTemplates";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Cargar Plantilla";
            this.gbTipo.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog ofdTemplate;
        private System.Windows.Forms.GroupBox gbTipo;
        private System.Windows.Forms.ComboBox cmbAnio;
        private System.Windows.Forms.ComboBox cmbTipoPlantilla;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCargar;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnSeleccionar;
        private System.Windows.Forms.TextBox txtPlantilla;
        private System.Windows.Forms.Label lblPlantilla;
    }
}