namespace ExcelAddIn1
{
    partial class FormulasComprobaciones
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormulasComprobaciones));
            this.ofdTemplate = new System.Windows.Forms.OpenFileDialog();
            this.gbProgress = new System.Windows.Forms.GroupBox();
            this.pgbFile = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.btnGenerar = new System.Windows.Forms.Button();
            this.fbdTemplate = new System.Windows.Forms.FolderBrowserDialog();
            this.pbTransferir = new System.Windows.Forms.PictureBox();
            this.gbProgress.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbTransferir)).BeginInit();
            this.SuspendLayout();
            // 
            // ofdTemplate
            // 
            this.ofdTemplate.Filter = "SAT Template | *.xlsm";
            // 
            // gbProgress
            // 
            this.gbProgress.Controls.Add(this.pgbFile);
            this.gbProgress.Location = new System.Drawing.Point(12, 35);
            this.gbProgress.Name = "gbProgress";
            this.gbProgress.Size = new System.Drawing.Size(426, 55);
            this.gbProgress.TabIndex = 9;
            this.gbProgress.TabStop = false;
            this.gbProgress.Text = "Progreso";
            // 
            // pgbFile
            // 
            this.pgbFile.Location = new System.Drawing.Point(6, 19);
            this.pgbFile.Name = "pgbFile";
            this.pgbFile.Size = new System.Drawing.Size(414, 23);
            this.pgbFile.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(426, 23);
            this.label1.TabIndex = 8;
            this.label1.Text = "Generando las formulas de Comprobaciones... Espere mientras termina el proceso.";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // btnGenerar
            // 
            this.btnGenerar.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("btnGenerar.BackgroundImage")));
            this.btnGenerar.Enabled = false;
            this.btnGenerar.Location = new System.Drawing.Point(357, 96);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Size = new System.Drawing.Size(75, 23);
            this.btnGenerar.TabIndex = 14;
            this.btnGenerar.Text = "  Aceptar";
            this.btnGenerar.UseVisualStyleBackColor = true;
            this.btnGenerar.Visible = false;
            this.btnGenerar.Click += new System.EventHandler(this.btnGenerar_Click);
            // 
            // pbTransferir
            // 
            this.pbTransferir.Image = ((System.Drawing.Image)(resources.GetObject("pbTransferir.Image")));
            this.pbTransferir.Location = new System.Drawing.Point(185, 44);
            this.pbTransferir.Name = "pbTransferir";
            this.pbTransferir.Size = new System.Drawing.Size(80, 50);
            this.pbTransferir.TabIndex = 16;
            this.pbTransferir.TabStop = false;
            // 
            // FormulasComprobaciones
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(451, 97);
            this.ControlBox = false;
            this.Controls.Add(this.pbTransferir);
            this.Controls.Add(this.gbProgress);
            this.Controls.Add(this.btnGenerar);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormulasComprobaciones";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Archivos Base";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.FormulasComprobaciones_Load);
            this.Shown += new System.EventHandler(this.FormulasComprobaciones_Shown);
            this.gbProgress.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbTransferir)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog ofdTemplate;
        private System.Windows.Forms.GroupBox gbProgress;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnGenerar;
        private System.Windows.Forms.ProgressBar pgbFile;
        private System.Windows.Forms.FolderBrowserDialog fbdTemplate;
        private System.Windows.Forms.PictureBox pbTransferir;
    }
}