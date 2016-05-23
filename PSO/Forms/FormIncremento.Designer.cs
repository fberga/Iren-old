namespace Iren.PSO.Forms
{
    partial class FormIncremento
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
            this.lbRangeSelezionato = new System.Windows.Forms.Label();
            this.chkTuttaRiga = new System.Windows.Forms.CheckBox();
            this.panelCentrale = new System.Windows.Forms.Panel();
            this.txtValore = new System.Windows.Forms.TextBox();
            this.txtPercentuale = new System.Windows.Forms.TextBox();
            this.rdbIncremento = new System.Windows.Forms.RadioButton();
            this.rdbPercentuale = new System.Windows.Forms.RadioButton();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.panelTop = new System.Windows.Forms.Panel();
            this.txtRangeSelezionato = new System.Windows.Forms.TextBox();
            this.lbScegli = new System.Windows.Forms.Label();
            this.panelCentrale.SuspendLayout();
            this.panelButtons.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbRangeSelezionato
            // 
            this.lbRangeSelezionato.AutoSize = true;
            this.lbRangeSelezionato.Location = new System.Drawing.Point(13, 12);
            this.lbRangeSelezionato.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbRangeSelezionato.Name = "lbRangeSelezionato";
            this.lbRangeSelezionato.Size = new System.Drawing.Size(149, 20);
            this.lbRangeSelezionato.TabIndex = 0;
            this.lbRangeSelezionato.Text = "Range Selezionato:";
            // 
            // chkTuttaRiga
            // 
            this.chkTuttaRiga.AutoSize = true;
            this.chkTuttaRiga.Location = new System.Drawing.Point(372, 11);
            this.chkTuttaRiga.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.chkTuttaRiga.Name = "chkTuttaRiga";
            this.chkTuttaRiga.Size = new System.Drawing.Size(111, 24);
            this.chkTuttaRiga.TabIndex = 2;
            this.chkTuttaRiga.Text = "Tutta la riga";
            this.chkTuttaRiga.UseVisualStyleBackColor = true;
            this.chkTuttaRiga.CheckedChanged += new System.EventHandler(this.chkTuttaRiga_CheckedChanged);
            // 
            // panelCentrale
            // 
            this.panelCentrale.Controls.Add(this.lbScegli);
            this.panelCentrale.Controls.Add(this.txtValore);
            this.panelCentrale.Controls.Add(this.txtPercentuale);
            this.panelCentrale.Controls.Add(this.rdbIncremento);
            this.panelCentrale.Controls.Add(this.rdbPercentuale);
            this.panelCentrale.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelCentrale.Location = new System.Drawing.Point(0, 47);
            this.panelCentrale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelCentrale.Name = "panelCentrale";
            this.panelCentrale.Size = new System.Drawing.Size(496, 112);
            this.panelCentrale.TabIndex = 5;
            // 
            // txtValore
            // 
            this.txtValore.Location = new System.Drawing.Point(166, 68);
            this.txtValore.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtValore.Name = "txtValore";
            this.txtValore.Size = new System.Drawing.Size(188, 26);
            this.txtValore.TabIndex = 7;
            // 
            // txtPercentuale
            // 
            this.txtPercentuale.Location = new System.Drawing.Point(166, 36);
            this.txtPercentuale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPercentuale.Name = "txtPercentuale";
            this.txtPercentuale.Size = new System.Drawing.Size(188, 26);
            this.txtPercentuale.TabIndex = 6;
            // 
            // rdbIncremento
            // 
            this.rdbIncremento.AutoSize = true;
            this.rdbIncremento.Location = new System.Drawing.Point(30, 70);
            this.rdbIncremento.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rdbIncremento.Name = "rdbIncremento";
            this.rdbIncremento.Size = new System.Drawing.Size(77, 24);
            this.rdbIncremento.TabIndex = 5;
            this.rdbIncremento.TabStop = true;
            this.rdbIncremento.Text = "Valore:";
            this.rdbIncremento.UseVisualStyleBackColor = true;
            // 
            // rdbPercentuale
            // 
            this.rdbPercentuale.AutoSize = true;
            this.rdbPercentuale.Location = new System.Drawing.Point(30, 36);
            this.rdbPercentuale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rdbPercentuale.Name = "rdbPercentuale";
            this.rdbPercentuale.Size = new System.Drawing.Size(116, 24);
            this.rdbPercentuale.TabIndex = 4;
            this.rdbPercentuale.TabStop = true;
            this.rdbPercentuale.Text = "Percentuale:";
            this.rdbPercentuale.UseVisualStyleBackColor = true;
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnApplica);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(0, 159);
            this.panelButtons.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 5, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(496, 53);
            this.panelButtons.TabIndex = 13;
            // 
            // btnApplica
            // 
            this.btnApplica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApplica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplica.Location = new System.Drawing.Point(270, 5);
            this.btnApplica.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnApplica.Name = "btnApplica";
            this.btnApplica.Size = new System.Drawing.Size(113, 48);
            this.btnApplica.TabIndex = 4;
            this.btnApplica.Text = "Modifica";
            this.btnApplica.UseVisualStyleBackColor = true;
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(383, 5);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 48);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.txtRangeSelezionato);
            this.panelTop.Controls.Add(this.lbRangeSelezionato);
            this.panelTop.Controls.Add(this.chkTuttaRiga);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(496, 47);
            this.panelTop.TabIndex = 14;
            // 
            // txtRangeSelezionato
            // 
            this.txtRangeSelezionato.Location = new System.Drawing.Point(169, 9);
            this.txtRangeSelezionato.Name = "txtRangeSelezionato";
            this.txtRangeSelezionato.Size = new System.Drawing.Size(185, 26);
            this.txtRangeSelezionato.TabIndex = 3;
            // 
            // lbScegli
            // 
            this.lbScegli.AutoSize = true;
            this.lbScegli.Location = new System.Drawing.Point(12, 5);
            this.lbScegli.Name = "lbScegli";
            this.lbScegli.Size = new System.Drawing.Size(195, 20);
            this.lbScegli.TabIndex = 8;
            this.lbScegli.Text = "Scegli il tipo di incremento:";
            // 
            // FormIncremento
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(496, 212);
            this.Controls.Add(this.panelCentrale);
            this.Controls.Add(this.panelTop);
            this.Controls.Add(this.panelButtons);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormIncremento";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormIncremento";
            this.panelCentrale.ResumeLayout(false);
            this.panelCentrale.PerformLayout();
            this.panelButtons.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lbRangeSelezionato;
        private System.Windows.Forms.CheckBox chkTuttaRiga;
        private System.Windows.Forms.Panel panelCentrale;
        private System.Windows.Forms.TextBox txtValore;
        private System.Windows.Forms.TextBox txtPercentuale;
        private System.Windows.Forms.RadioButton rdbIncremento;
        private System.Windows.Forms.RadioButton rdbPercentuale;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnApplica;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.TextBox txtRangeSelezionato;
        private System.Windows.Forms.Label lbScegli;
    }
}