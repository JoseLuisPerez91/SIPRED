namespace ExcelAddIn1 {
    partial class Cruce {
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
            this.lblTitle = new System.Windows.Forms.Label();
            this.gbTipo = new System.Windows.Forms.GroupBox();
            this.cmbTipo = new System.Windows.Forms.ComboBox();
            this.ckbValidacionMinima = new System.Windows.Forms.CheckBox();
            this.ckbValidarCalculos = new System.Windows.Forms.CheckBox();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnAceptar = new System.Windows.Forms.Button();
            this.gbTipo.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.BackColor = System.Drawing.SystemColors.Info;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(13, 13);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(514, 23);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "¿Desea realizar el proceso de verificación de cruces?";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // gbTipo
            // 
            this.gbTipo.Controls.Add(this.cmbTipo);
            this.gbTipo.Location = new System.Drawing.Point(13, 45);
            this.gbTipo.Name = "gbTipo";
            this.gbTipo.Size = new System.Drawing.Size(514, 54);
            this.gbTipo.TabIndex = 1;
            this.gbTipo.TabStop = false;
            this.gbTipo.Text = "Tipo";
            // 
            // cmbTipo
            // 
            this.cmbTipo.FormattingEnabled = true;
            this.cmbTipo.Location = new System.Drawing.Point(7, 20);
            this.cmbTipo.Name = "cmbTipo";
            this.cmbTipo.Size = new System.Drawing.Size(501, 21);
            this.cmbTipo.TabIndex = 0;
            // 
            // ckbValidacionMinima
            // 
            this.ckbValidacionMinima.AutoSize = true;
            this.ckbValidacionMinima.Location = new System.Drawing.Point(13, 106);
            this.ckbValidacionMinima.Name = "ckbValidacionMinima";
            this.ckbValidacionMinima.Size = new System.Drawing.Size(418, 17);
            this.ckbValidacionMinima.TabIndex = 2;
            this.ckbValidacionMinima.Text = "Validar información minima, sin signo, excluyente, cuestionarios y otras validaci" +
    "ones";
            this.ckbValidacionMinima.UseVisualStyleBackColor = true;
            // 
            // ckbValidarCalculos
            // 
            this.ckbValidarCalculos.AutoSize = true;
            this.ckbValidarCalculos.Location = new System.Drawing.Point(13, 130);
            this.ckbValidarCalculos.Name = "ckbValidarCalculos";
            this.ckbValidarCalculos.Size = new System.Drawing.Size(295, 17);
            this.ckbValidarCalculos.TabIndex = 3;
            this.ckbValidarCalculos.Text = "Validar cálculos de fórmulas de SIPRED, SIPIAD o DISIF";
            this.ckbValidarCalculos.UseVisualStyleBackColor = true;
            // 
            // btnCancelar
            // 
            this.btnCancelar.Location = new System.Drawing.Point(271, 156);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(75, 23);
            this.btnCancelar.TabIndex = 4;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            // 
            // btnAceptar
            // 
            this.btnAceptar.Location = new System.Drawing.Point(190, 156);
            this.btnAceptar.Name = "btnAceptar";
            this.btnAceptar.Size = new System.Drawing.Size(75, 23);
            this.btnAceptar.TabIndex = 5;
            this.btnAceptar.Text = "Aceptar";
            this.btnAceptar.UseVisualStyleBackColor = true;
            this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
            // 
            // Cruce
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(539, 190);
            this.Controls.Add(this.btnAceptar);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.ckbValidarCalculos);
            this.Controls.Add(this.ckbValidacionMinima);
            this.Controls.Add(this.gbTipo);
            this.Controls.Add(this.lblTitle);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Cruce";
            this.Text = "Verificación de Cruces";
            this.gbTipo.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.GroupBox gbTipo;
        private System.Windows.Forms.ComboBox cmbTipo;
        private System.Windows.Forms.CheckBox ckbValidacionMinima;
        private System.Windows.Forms.CheckBox ckbValidarCalculos;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnAceptar;
    }
}