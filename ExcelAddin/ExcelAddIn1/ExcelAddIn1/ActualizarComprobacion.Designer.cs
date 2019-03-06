namespace ExcelAddIn1
{
    partial class ActualizarComprobacion
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ActualizarComprobacion));
            this.tbcCruce = new System.Windows.Forms.TabControl();
            this.tbDefinicion = new System.Windows.Forms.TabPage();
            this.txtformula = new System.Windows.Forms.TextBox();
            this.txtcelda = new System.Windows.Forms.TextBox();
            this.txtConcepto = new System.Windows.Forms.TextBox();
            this.gbCondicionar = new System.Windows.Forms.GroupBox();
            this.chkCondicionar = new System.Windows.Forms.CheckBox();
            this.txtCondicion = new System.Windows.Forms.TextBox();
            this.lblcontra = new System.Windows.Forms.Label();
            this.lblcruzar = new System.Windows.Forms.Label();
            this.lblconcepto = new System.Windows.Forms.Label();
            this.txtNro = new System.Windows.Forms.TextBox();
            this.lblnro = new System.Windows.Forms.Label();
            this.tbNota = new System.Windows.Forms.TabPage();
            this.txtNota = new System.Windows.Forms.RichTextBox();
            this.btncancelar = new System.Windows.Forms.Button();
            this.btguardar = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tbcCruce.SuspendLayout();
            this.tbDefinicion.SuspendLayout();
            this.gbCondicionar.SuspendLayout();
            this.tbNota.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tbcCruce
            // 
            this.tbcCruce.Controls.Add(this.tbDefinicion);
            this.tbcCruce.Controls.Add(this.tbNota);
            this.tbcCruce.Location = new System.Drawing.Point(1, 2);
            this.tbcCruce.Name = "tbcCruce";
            this.tbcCruce.SelectedIndex = 0;
            this.tbcCruce.Size = new System.Drawing.Size(545, 343);
            this.tbcCruce.TabIndex = 1;
            // 
            // tbDefinicion
            // 
            this.tbDefinicion.Controls.Add(this.txtformula);
            this.tbDefinicion.Controls.Add(this.txtcelda);
            this.tbDefinicion.Controls.Add(this.txtConcepto);
            this.tbDefinicion.Controls.Add(this.gbCondicionar);
            this.tbDefinicion.Controls.Add(this.lblcontra);
            this.tbDefinicion.Controls.Add(this.lblcruzar);
            this.tbDefinicion.Controls.Add(this.lblconcepto);
            this.tbDefinicion.Controls.Add(this.txtNro);
            this.tbDefinicion.Controls.Add(this.lblnro);
            this.tbDefinicion.Location = new System.Drawing.Point(4, 22);
            this.tbDefinicion.Name = "tbDefinicion";
            this.tbDefinicion.Padding = new System.Windows.Forms.Padding(3);
            this.tbDefinicion.Size = new System.Drawing.Size(537, 317);
            this.tbDefinicion.TabIndex = 0;
            this.tbDefinicion.Text = "Definición";
            this.tbDefinicion.UseVisualStyleBackColor = true;
            // 
            // txtformula
            // 
            this.txtformula.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtformula.Location = new System.Drawing.Point(87, 131);
            this.txtformula.Name = "txtformula";
            this.txtformula.Size = new System.Drawing.Size(437, 20);
            this.txtformula.TabIndex = 4;
            // 
            // txtcelda
            // 
            this.txtcelda.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtcelda.Location = new System.Drawing.Point(87, 96);
            this.txtcelda.Name = "txtcelda";
            this.txtcelda.Size = new System.Drawing.Size(437, 20);
            this.txtcelda.TabIndex = 3;
            // 
            // txtConcepto
            // 
            this.txtConcepto.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtConcepto.Location = new System.Drawing.Point(87, 61);
            this.txtConcepto.Name = "txtConcepto";
            this.txtConcepto.Size = new System.Drawing.Size(437, 20);
            this.txtConcepto.TabIndex = 2;
            // 
            // gbCondicionar
            // 
            this.gbCondicionar.Controls.Add(this.chkCondicionar);
            this.gbCondicionar.Controls.Add(this.txtCondicion);
            this.gbCondicionar.Location = new System.Drawing.Point(6, 167);
            this.gbCondicionar.Name = "gbCondicionar";
            this.gbCondicionar.Size = new System.Drawing.Size(528, 56);
            this.gbCondicionar.TabIndex = 5;
            this.gbCondicionar.TabStop = false;
            // 
            // chkCondicionar
            // 
            this.chkCondicionar.AutoSize = true;
            this.chkCondicionar.Location = new System.Drawing.Point(14, 0);
            this.chkCondicionar.Name = "chkCondicionar";
            this.chkCondicionar.Size = new System.Drawing.Size(82, 17);
            this.chkCondicionar.TabIndex = 5;
            this.chkCondicionar.Text = "Condicionar";
            this.chkCondicionar.UseVisualStyleBackColor = true;
            this.chkCondicionar.CheckedChanged += new System.EventHandler(this.chkCondicionar_CheckedChanged);
            // 
            // txtCondicion
            // 
            this.txtCondicion.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.txtCondicion.Location = new System.Drawing.Point(14, 23);
            this.txtCondicion.Name = "txtCondicion";
            this.txtCondicion.ReadOnly = true;
            this.txtCondicion.Size = new System.Drawing.Size(504, 20);
            this.txtCondicion.TabIndex = 0;
            // 
            // lblcontra
            // 
            this.lblcontra.AutoSize = true;
            this.lblcontra.Location = new System.Drawing.Point(17, 134);
            this.lblcontra.Name = "lblcontra";
            this.lblcontra.Size = new System.Drawing.Size(44, 13);
            this.lblcontra.TabIndex = 4;
            this.lblcontra.Text = "Fórmula";
            // 
            // lblcruzar
            // 
            this.lblcruzar.AutoSize = true;
            this.lblcruzar.Location = new System.Drawing.Point(20, 96);
            this.lblcruzar.Name = "lblcruzar";
            this.lblcruzar.Size = new System.Drawing.Size(34, 13);
            this.lblcruzar.TabIndex = 3;
            this.lblcruzar.Text = "Celda";
            // 
            // lblconcepto
            // 
            this.lblconcepto.AutoSize = true;
            this.lblconcepto.Location = new System.Drawing.Point(17, 61);
            this.lblconcepto.Name = "lblconcepto";
            this.lblconcepto.Size = new System.Drawing.Size(53, 13);
            this.lblconcepto.TabIndex = 2;
            this.lblconcepto.Text = "Concepto";
            // 
            // txtNro
            // 
            this.txtNro.Location = new System.Drawing.Point(87, 25);
            this.txtNro.Name = "txtNro";
            this.txtNro.ReadOnly = true;
            this.txtNro.Size = new System.Drawing.Size(100, 20);
            this.txtNro.TabIndex = 1;
            // 
            // lblnro
            // 
            this.lblnro.AutoSize = true;
            this.lblnro.Location = new System.Drawing.Point(17, 25);
            this.lblnro.Name = "lblnro";
            this.lblnro.Size = new System.Drawing.Size(44, 13);
            this.lblnro.TabIndex = 0;
            this.lblnro.Text = "Número";
            // 
            // tbNota
            // 
            this.tbNota.Controls.Add(this.txtNota);
            this.tbNota.Location = new System.Drawing.Point(4, 22);
            this.tbNota.Name = "tbNota";
            this.tbNota.Padding = new System.Windows.Forms.Padding(3);
            this.tbNota.Size = new System.Drawing.Size(537, 317);
            this.tbNota.TabIndex = 1;
            this.tbNota.Text = "Nota";
            this.tbNota.UseVisualStyleBackColor = true;
            // 
            // txtNota
            // 
            this.txtNota.BackColor = System.Drawing.SystemColors.Info;
            this.txtNota.Location = new System.Drawing.Point(0, 3);
            this.txtNota.Name = "txtNota";
            this.txtNota.Size = new System.Drawing.Size(534, 314);
            this.txtNota.TabIndex = 8;
            this.txtNota.Text = "";
            this.txtNota.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNota_KeyPress);
            // 
            // btncancelar
            // 
            this.btncancelar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btncancelar.BackgroundImage")));
            this.btncancelar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.btncancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btncancelar.Location = new System.Drawing.Point(92, 3);
            this.btncancelar.Name = "btncancelar";
            this.btncancelar.Size = new System.Drawing.Size(83, 29);
            this.btncancelar.TabIndex = 8;
            this.btncancelar.Text = "  Cancelar";
            this.btncancelar.UseVisualStyleBackColor = true;
            // 
            // btguardar
            // 
            this.btguardar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btguardar.BackgroundImage")));
            this.btguardar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.btguardar.Location = new System.Drawing.Point(3, 3);
            this.btguardar.Name = "btguardar";
            this.btguardar.Size = new System.Drawing.Size(83, 29);
            this.btguardar.TabIndex = 9;
            this.btguardar.Text = "  Guardar";
            this.btguardar.UseVisualStyleBackColor = true;
            this.btguardar.Click += new System.EventHandler(this.btguardar_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btguardar);
            this.panel1.Controls.Add(this.btncancelar);
            this.panel1.Location = new System.Drawing.Point(361, 351);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(181, 37);
            this.panel1.TabIndex = 10;
            // 
            // ActualizarComprobacion
            // 
            this.AcceptButton = this.btguardar;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(543, 392);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.tbcCruce);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "ActualizarComprobacion";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ActualizarComprobacion";
            this.TopMost = true;
            this.tbcCruce.ResumeLayout(false);
            this.tbDefinicion.ResumeLayout(false);
            this.tbDefinicion.PerformLayout();
            this.gbCondicionar.ResumeLayout(false);
            this.gbCondicionar.PerformLayout();
            this.tbNota.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tbcCruce;
        private System.Windows.Forms.TabPage tbDefinicion;
        private System.Windows.Forms.TextBox txtformula;
        private System.Windows.Forms.TextBox txtcelda;
        private System.Windows.Forms.TextBox txtConcepto;
        private System.Windows.Forms.GroupBox gbCondicionar;
        private System.Windows.Forms.CheckBox chkCondicionar;
        private System.Windows.Forms.TextBox txtCondicion;
        private System.Windows.Forms.Label lblcontra;
        private System.Windows.Forms.Label lblcruzar;
        private System.Windows.Forms.Label lblconcepto;
        private System.Windows.Forms.TextBox txtNro;
        private System.Windows.Forms.Label lblnro;
        private System.Windows.Forms.TabPage tbNota;
        private System.Windows.Forms.RichTextBox txtNota;
        private System.Windows.Forms.Button btncancelar;
        private System.Windows.Forms.Button btguardar;
        private System.Windows.Forms.Panel panel1;
    }
}