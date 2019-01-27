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
            this.ckbValidacionMinima = new System.Windows.Forms.CheckBox();
            this.ckbValidarCalculos = new System.Windows.Forms.CheckBox();
            this.btnCancelar = new System.Windows.Forms.Button();
            this.btnAceptar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.BackColor = System.Drawing.SystemColors.Info;
            this.lblTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(17, 16);
            this.lblTitle.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(685, 28);
            this.lblTitle.TabIndex = 0;
            this.lblTitle.Text = "¿Desea realizar el proceso de verificación de cruces?";
            this.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ckbValidacionMinima
            // 
            this.ckbValidacionMinima.AutoSize = true;
            this.ckbValidacionMinima.Location = new System.Drawing.Point(16, 59);
            this.ckbValidacionMinima.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ckbValidacionMinima.Name = "ckbValidacionMinima";
            this.ckbValidacionMinima.Size = new System.Drawing.Size(559, 21);
            this.ckbValidacionMinima.TabIndex = 2;
            this.ckbValidacionMinima.Text = "Validar información minima, sin signo, excluyente, cuestionarios y otras validaci" +
    "ones";
            this.ckbValidacionMinima.UseVisualStyleBackColor = true;
            // 
            // ckbValidarCalculos
            // 
            this.ckbValidarCalculos.AutoSize = true;
            this.ckbValidarCalculos.Location = new System.Drawing.Point(16, 87);
            this.ckbValidarCalculos.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ckbValidarCalculos.Name = "ckbValidarCalculos";
            this.ckbValidarCalculos.Size = new System.Drawing.Size(381, 21);
            this.ckbValidarCalculos.TabIndex = 3;
            this.ckbValidarCalculos.Text = "Validar cálculos de fórmulas de SIPRED, SIPIAD o DISIF";
            this.ckbValidarCalculos.UseVisualStyleBackColor = true;
            // 
            // btnCancelar
            // 
            this.btnCancelar.Location = new System.Drawing.Point(361, 118);
            this.btnCancelar.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCancelar.Name = "btnCancelar";
            this.btnCancelar.Size = new System.Drawing.Size(100, 28);
            this.btnCancelar.TabIndex = 4;
            this.btnCancelar.Text = "Cancelar";
            this.btnCancelar.UseVisualStyleBackColor = true;
            this.btnCancelar.Click += new System.EventHandler(this.btnCancelar_Click);
            // 
            // btnAceptar
            // 
            this.btnAceptar.Location = new System.Drawing.Point(253, 118);
            this.btnAceptar.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAceptar.Name = "btnAceptar";
            this.btnAceptar.Size = new System.Drawing.Size(100, 28);
            this.btnAceptar.TabIndex = 5;
            this.btnAceptar.Text = "Aceptar";
            this.btnAceptar.UseVisualStyleBackColor = true;
            this.btnAceptar.Click += new System.EventHandler(this.btnAceptar_Click);
            // 
            // Cruce
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(719, 160);
            this.Controls.Add(this.btnAceptar);
            this.Controls.Add(this.btnCancelar);
            this.Controls.Add(this.ckbValidarCalculos);
            this.Controls.Add(this.ckbValidacionMinima);
            this.Controls.Add(this.lblTitle);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Cruce";
            this.Text = "Verificación de Cruces";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.CheckBox ckbValidacionMinima;
        private System.Windows.Forms.CheckBox ckbValidarCalculos;
        private System.Windows.Forms.Button btnCancelar;
        private System.Windows.Forms.Button btnAceptar;
    }
}