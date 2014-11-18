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
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cmbRichiesta
            // 
            this.cmbRichiesta.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbRichiesta.FormattingEnabled = true;
            this.cmbRichiesta.Location = new System.Drawing.Point(148, 29);
            this.cmbRichiesta.Name = "cmbRichiesta";
            this.cmbRichiesta.Size = new System.Drawing.Size(208, 28);
            this.cmbRichiesta.TabIndex = 0;
            // 
            // lbIdRichiesta
            // 
            this.lbIdRichiesta.AutoSize = true;
            this.lbIdRichiesta.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbIdRichiesta.Location = new System.Drawing.Point(12, 32);
            this.lbIdRichiesta.Name = "lbIdRichiesta";
            this.lbIdRichiesta.Size = new System.Drawing.Size(130, 25);
            this.lbIdRichiesta.TabIndex = 1;
            this.lbIdRichiesta.Text = "N° Richiesta";
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(223, 79);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(133, 47);
            this.btnAnnulla.TabIndex = 3;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);
            // 
            // btnOK
            // 
            this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.Location = new System.Drawing.Point(84, 79);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(133, 47);
            this.btnOK.TabIndex = 4;
            this.btnOK.Text = "Ok";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // FormAnnullaModifica
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnAnnulla;
            this.ClientSize = new System.Drawing.Size(368, 138);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnAnnulla);
            this.Controls.Add(this.lbIdRichiesta);
            this.Controls.Add(this.cmbRichiesta);
            this.Name = "FormAnnullaModifica";
            this.ShowIcon = false;
            this.Text = "Annulla Modifica";
            this.Load += new System.EventHandler(this.FormAnnullaModifica_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbRichiesta;
        private System.Windows.Forms.Label lbIdRichiesta;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Button btnOK;
    }
}