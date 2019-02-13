namespace ExcelAddIn1
{
    partial class Explicaciones
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
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnAccept = new System.Windows.Forms.Button();
            this.TxtExplicacion = new System.Windows.Forms.RichTextBox();
            this.lblcontador = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(372, 370);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(90, 23);
            this.btnCancel.TabIndex = 12;
            this.btnCancel.Text = "Cancelar";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnAccept
            // 
            this.btnAccept.Location = new System.Drawing.Point(263, 370);
            this.btnAccept.Name = "btnAccept";
            this.btnAccept.Size = new System.Drawing.Size(90, 23);
            this.btnAccept.TabIndex = 11;
            this.btnAccept.Text = "Aceptar";
            this.btnAccept.UseVisualStyleBackColor = true;
            this.btnAccept.Click += new System.EventHandler(this.btnAccept_Click);
            // 
            // TxtExplicacion
            // 
            this.TxtExplicacion.BackColor = System.Drawing.SystemColors.Info;
            this.TxtExplicacion.Font = new System.Drawing.Font("Arial", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TxtExplicacion.ForeColor = System.Drawing.SystemColors.InfoText;
            this.TxtExplicacion.Location = new System.Drawing.Point(4, 11);
            this.TxtExplicacion.Name = "TxtExplicacion";
            this.TxtExplicacion.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.ForcedVertical;
            this.TxtExplicacion.Size = new System.Drawing.Size(684, 353);
            this.TxtExplicacion.TabIndex = 10;
            this.TxtExplicacion.Text = "";
            this.TxtExplicacion.TextChanged += new System.EventHandler(this.TxtExplicacion_TextChanged);
            this.TxtExplicacion.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtExplicacion_KeyPress);
            // 
            // lblcontador
            // 
            this.lblcontador.AutoSize = true;
            this.lblcontador.Location = new System.Drawing.Point(12, 375);
            this.lblcontador.Name = "lblcontador";
            this.lblcontador.Size = new System.Drawing.Size(0, 13);
            this.lblcontador.TabIndex = 13;
            // 
            // Explicaciones
            // 
            this.AcceptButton = this.btnAccept;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(692, 404);
            this.Controls.Add(this.lblcontador);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnAccept);
            this.Controls.Add(this.TxtExplicacion);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Explicaciones";
            this.ShowIcon = false;
            this.Text = "x";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnAccept;
        private System.Windows.Forms.RichTextBox TxtExplicacion;
        private System.Windows.Forms.Label lblcontador;
    }
}