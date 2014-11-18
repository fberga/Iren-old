namespace RiMoST2
{
    partial class FormAnnullaModifica
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
            this.cmbRichiesta = new System.Windows.Forms.ComboBox();
            this.lbIdRichiesta = new System.Windows.Forms.Label();
            this.DocPreview = new System.Windows.Forms.WebBrowser();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cmbRichiesta
            // 
            this.cmbRichiesta.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbRichiesta.FormattingEnabled = true;
            this.cmbRichiesta.Location = new System.Drawing.Point(143, 29);
            this.cmbRichiesta.Name = "cmbRichiesta";
            this.cmbRichiesta.Size = new System.Drawing.Size(288, 28);
            this.cmbRichiesta.TabIndex = 0;
            this.cmbRichiesta.SelectedIndexChanged += new System.EventHandler(this.cmbRichiesta_SelectedIndexChanged);
            // 
            // lbIdRichiesta
            // 
            this.lbIdRichiesta.AutoSize = true;
            this.lbIdRichiesta.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbIdRichiesta.Location = new System.Drawing.Point(7, 32);
            this.lbIdRichiesta.Name = "lbIdRichiesta";
            this.lbIdRichiesta.Size = new System.Drawing.Size(130, 25);
            this.lbIdRichiesta.TabIndex = 1;
            this.lbIdRichiesta.Text = "N° Richiesta";
            // 
            // DocPreview
            // 
            this.DocPreview.Location = new System.Drawing.Point(12, 63);
            this.DocPreview.MinimumSize = new System.Drawing.Size(20, 20);
            this.DocPreview.Name = "DocPreview";
            this.DocPreview.ScrollBarsEnabled = false;
            this.DocPreview.Size = new System.Drawing.Size(713, 440);
            this.DocPreview.TabIndex = 5;
            // 
            // btnOK
            // 
            this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.Location = new System.Drawing.Point(453, 509);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(133, 47);
            this.btnOK.TabIndex = 6;
            this.btnOK.Text = "Ok";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(592, 509);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(133, 47);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);
            // 
            // FormAnnullaModifica
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(737, 568);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnAnnulla);
            this.Controls.Add(this.DocPreview);
            this.Controls.Add(this.lbIdRichiesta);
            this.Controls.Add(this.cmbRichiesta);
            this.Name = "FormAnnullaModifica";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Annulla Modifica";
            this.Load += new System.EventHandler(this.FormAnnullaModifica_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbRichiesta;
        private System.Windows.Forms.Label lbIdRichiesta;
        private System.Windows.Forms.WebBrowser DocPreview;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnAnnulla;
    }
}