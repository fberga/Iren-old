namespace Iren.FrontOffice.Forms
{
    partial class FormSelezioneUP
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
            this.comboUP = new System.Windows.Forms.ComboBox();
            this.lblSeleziona = new System.Windows.Forms.Label();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnCarica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.panelButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboUP
            // 
            this.comboUP.FormattingEnabled = true;
            this.comboUP.Location = new System.Drawing.Point(18, 34);
            this.comboUP.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.comboUP.Name = "comboUP";
            this.comboUP.Size = new System.Drawing.Size(369, 28);
            this.comboUP.TabIndex = 0;
            this.comboUP.SelectedIndexChanged += new System.EventHandler(this.comboUP_SelectedIndexChanged);
            // 
            // lblSeleziona
            // 
            this.lblSeleziona.AutoSize = true;
            this.lblSeleziona.Location = new System.Drawing.Point(17, 12);
            this.lblSeleziona.Name = "lblSeleziona";
            this.lblSeleziona.Size = new System.Drawing.Size(272, 20);
            this.lblSeleziona.TabIndex = 1;
            this.lblSeleziona.Text = "Selezionare l\'elemento da ottimizzare";
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnCarica);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(3, 69);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(394, 53);
            this.panelButtons.TabIndex = 14;
            // 
            // btnCarica
            // 
            this.btnCarica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnCarica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCarica.Location = new System.Drawing.Point(168, 3);
            this.btnCarica.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCarica.Name = "btnCarica";
            this.btnCarica.Size = new System.Drawing.Size(113, 50);
            this.btnCarica.TabIndex = 4;
            this.btnCarica.Text = "Seleziona";
            this.btnCarica.UseVisualStyleBackColor = true;
            this.btnCarica.Click += new System.EventHandler(this.btnCarica_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(281, 3);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 50);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);
            // 
            // frmSELUP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 125);
            this.Controls.Add(this.panelButtons);
            this.Controls.Add(this.lblSeleziona);
            this.Controls.Add(this.comboUP);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "frmSELUP";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmSELUP";
            this.Load += new System.EventHandler(this.frmSELUP_Load);
            this.panelButtons.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboUP;
        private System.Windows.Forms.Label lblSeleziona;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnCarica;
        private System.Windows.Forms.Button btnAnnulla;
    }
}