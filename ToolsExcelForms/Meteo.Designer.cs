namespace Iren.FrontOffice.Forms
{
    partial class frmMETEO
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
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnCarica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.groupDati = new System.Windows.Forms.GroupBox();
            this.panelTop = new System.Windows.Forms.Panel();
            this.labelData = new System.Windows.Forms.Label();
            this.labelDataEmissione = new System.Windows.Forms.Label();
            this.radioNIMBUS = new System.Windows.Forms.RadioButton();
            this.comboNIMBUS = new System.Windows.Forms.ComboBox();
            this.radioEPSON = new System.Windows.Forms.RadioButton();
            this.comboEPSON = new System.Windows.Forms.ComboBox();
            this.radioARPA = new System.Windows.Forms.RadioButton();
            this.comboARPA = new System.Windows.Forms.ComboBox();
            this.panelButtons.SuspendLayout();
            this.groupDati.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // comboUP
            // 
            this.comboUP.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.comboUP.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboUP.FormattingEnabled = true;
            this.comboUP.Location = new System.Drawing.Point(5, 34);
            this.comboUP.Name = "comboUP";
            this.comboUP.Size = new System.Drawing.Size(331, 24);
            this.comboUP.TabIndex = 0;
            this.comboUP.SelectedIndexChanged += new System.EventHandler(this.comboUP_SelectedIndexChanged);
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnCarica);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(5, 215);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(341, 53);
            this.panelButtons.TabIndex = 13;
            // 
            // btnCarica
            // 
            this.btnCarica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnCarica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCarica.Location = new System.Drawing.Point(115, 3);
            this.btnCarica.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnCarica.Name = "btnCarica";
            this.btnCarica.Size = new System.Drawing.Size(113, 50);
            this.btnCarica.TabIndex = 4;
            this.btnCarica.Text = "Carica";
            this.btnCarica.UseVisualStyleBackColor = true;
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(228, 3);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 50);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // groupDati
            // 
            this.groupDati.Controls.Add(this.comboNIMBUS);
            this.groupDati.Controls.Add(this.comboEPSON);
            this.groupDati.Controls.Add(this.comboARPA);
            this.groupDati.Controls.Add(this.radioNIMBUS);
            this.groupDati.Controls.Add(this.radioEPSON);
            this.groupDati.Controls.Add(this.radioARPA);
            this.groupDati.Controls.Add(this.labelDataEmissione);
            this.groupDati.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupDati.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupDati.Location = new System.Drawing.Point(5, 68);
            this.groupDati.Name = "groupDati";
            this.groupDati.Size = new System.Drawing.Size(341, 147);
            this.groupDati.TabIndex = 14;
            this.groupDati.TabStop = false;
            this.groupDati.Text = "Fonti Previsione Meteo";
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.comboUP);
            this.panelTop.Controls.Add(this.labelData);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(5, 5);
            this.panelTop.Name = "panelTop";
            this.panelTop.Padding = new System.Windows.Forms.Padding(5);
            this.panelTop.Size = new System.Drawing.Size(341, 63);
            this.panelTop.TabIndex = 7;
            // 
            // labelData
            // 
            this.labelData.AutoSize = true;
            this.labelData.Dock = System.Windows.Forms.DockStyle.Top;
            this.labelData.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelData.Location = new System.Drawing.Point(5, 5);
            this.labelData.Name = "labelData";
            this.labelData.Size = new System.Drawing.Size(127, 18);
            this.labelData.TabIndex = 1;
            this.labelData.Text = "Data Riferimento: ";
            // 
            // labelDataEmissione
            // 
            this.labelDataEmissione.AutoSize = true;
            this.labelDataEmissione.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelDataEmissione.Location = new System.Drawing.Point(115, 26);
            this.labelDataEmissione.Name = "labelDataEmissione";
            this.labelDataEmissione.Size = new System.Drawing.Size(113, 18);
            this.labelDataEmissione.TabIndex = 0;
            this.labelDataEmissione.Text = "Data Emissione";
            // 
            // radioNIMBUS
            // 
            this.radioNIMBUS.AutoSize = true;
            this.radioNIMBUS.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioNIMBUS.Location = new System.Drawing.Point(5, 108);
            this.radioNIMBUS.Name = "radioNIMBUS";
            this.radioNIMBUS.Size = new System.Drawing.Size(90, 24);
            this.radioNIMBUS.TabIndex = 3;
            this.radioNIMBUS.TabStop = true;
            this.radioNIMBUS.Text = "NIMBUS";
            this.radioNIMBUS.UseVisualStyleBackColor = true;
            // 
            // comboNIMBUS
            // 
            this.comboNIMBUS.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboNIMBUS.FormattingEnabled = true;
            this.comboNIMBUS.Location = new System.Drawing.Point(116, 106);
            this.comboNIMBUS.Name = "comboNIMBUS";
            this.comboNIMBUS.Size = new System.Drawing.Size(220, 28);
            this.comboNIMBUS.TabIndex = 6;
            this.comboNIMBUS.DataSourceChanged += new System.EventHandler(this.comboDataEmissione_DataSourceChanged);
            // 
            // radioEPSON
            // 
            this.radioEPSON.AutoSize = true;
            this.radioEPSON.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioEPSON.Location = new System.Drawing.Point(8, 78);
            this.radioEPSON.Name = "radioEPSON";
            this.radioEPSON.Size = new System.Drawing.Size(82, 24);
            this.radioEPSON.TabIndex = 2;
            this.radioEPSON.TabStop = true;
            this.radioEPSON.Text = "EPSON";
            this.radioEPSON.UseVisualStyleBackColor = true;
            // 
            // comboEPSON
            // 
            this.comboEPSON.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboEPSON.FormattingEnabled = true;
            this.comboEPSON.Location = new System.Drawing.Point(116, 76);
            this.comboEPSON.Name = "comboEPSON";
            this.comboEPSON.Size = new System.Drawing.Size(220, 28);
            this.comboEPSON.TabIndex = 5;
            this.comboEPSON.DataSourceChanged += new System.EventHandler(this.comboDataEmissione_DataSourceChanged);
            // 
            // radioARPA
            // 
            this.radioARPA.AutoSize = true;
            this.radioARPA.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.radioARPA.Location = new System.Drawing.Point(8, 48);
            this.radioARPA.Name = "radioARPA";
            this.radioARPA.Size = new System.Drawing.Size(71, 24);
            this.radioARPA.TabIndex = 1;
            this.radioARPA.TabStop = true;
            this.radioARPA.Text = "ARPA";
            this.radioARPA.UseVisualStyleBackColor = true;
            // 
            // comboARPA
            // 
            this.comboARPA.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboARPA.FormattingEnabled = true;
            this.comboARPA.Location = new System.Drawing.Point(116, 46);
            this.comboARPA.Name = "comboARPA";
            this.comboARPA.Size = new System.Drawing.Size(220, 28);
            this.comboARPA.TabIndex = 4;
            this.comboARPA.DataSourceChanged += new System.EventHandler(this.comboDataEmissione_DataSourceChanged);
            // 
            // frmMETEO
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(351, 273);
            this.Controls.Add(this.groupDati);
            this.Controls.Add(this.panelTop);
            this.Controls.Add(this.panelButtons);
            this.Name = "frmMETEO";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.Text = "Meteo";
            this.Load += new System.EventHandler(this.frmMETEO_Load);
            this.panelButtons.ResumeLayout(false);
            this.groupDati.ResumeLayout(false);
            this.groupDati.PerformLayout();
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox comboUP;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnCarica;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.GroupBox groupDati;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.Label labelData;
        private System.Windows.Forms.Label labelDataEmissione;
        private System.Windows.Forms.ComboBox comboNIMBUS;
        private System.Windows.Forms.ComboBox comboEPSON;
        private System.Windows.Forms.ComboBox comboARPA;
        private System.Windows.Forms.RadioButton radioNIMBUS;
        private System.Windows.Forms.RadioButton radioEPSON;
        private System.Windows.Forms.RadioButton radioARPA;
    }
}