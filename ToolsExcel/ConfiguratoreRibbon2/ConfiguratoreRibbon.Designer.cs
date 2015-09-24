namespace ConfiguratoreRibbon2
{
    partial class ConfiguratoreRibbon
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConfiguratoreRibbon));
            this.toolStripTopMenu = new System.Windows.Forms.ToolStrip();
            this.AddGroup = new System.Windows.Forms.ToolStripButton();
            this.AddButton = new System.Windows.Forms.ToolStripButton();
            this.panelRibbonLayout = new System.Windows.Forms.Panel();
            this.panelFill = new System.Windows.Forms.Panel();
            this.chooseImageDialog = new System.Windows.Forms.OpenFileDialog();
            this.toolStripTopMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripTopMenu
            // 
            this.toolStripTopMenu.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripTopMenu.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.toolStripTopMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddGroup,
            this.AddButton});
            this.toolStripTopMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripTopMenu.Name = "toolStripTopMenu";
            this.toolStripTopMenu.Size = new System.Drawing.Size(1302, 54);
            this.toolStripTopMenu.TabIndex = 2;
            this.toolStripTopMenu.Text = "toolStrip1";
            // 
            // AddGroup
            // 
            this.AddGroup.Image = global::ConfiguratoreRibbon2.Properties.Resources.add_icon;
            this.AddGroup.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.AddGroup.Name = "AddGroup";
            this.AddGroup.Size = new System.Drawing.Size(51, 51);
            this.AddGroup.Text = "Gruppo";
            this.AddGroup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.AddGroup.Click += new System.EventHandler(this.AggiungiGruppo_Click);
            // 
            // AddButton
            // 
            this.AddButton.Image = ((System.Drawing.Image)(resources.GetObject("AddButton.Image")));
            this.AddButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.AddButton.Name = "AddButton";
            this.AddButton.Size = new System.Drawing.Size(40, 51);
            this.AddButton.Text = "Tasto";
            this.AddButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.AddButton.Click += new System.EventHandler(this.AggiungiTasto_Click);
            // 
            // panelRibbonLayout
            // 
            this.panelRibbonLayout.BackColor = System.Drawing.SystemColors.ControlLight;
            this.panelRibbonLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelRibbonLayout.Location = new System.Drawing.Point(0, 54);
            this.panelRibbonLayout.Margin = new System.Windows.Forms.Padding(4);
            this.panelRibbonLayout.Name = "panelRibbonLayout";
            this.panelRibbonLayout.Padding = new System.Windows.Forms.Padding(4);
            this.panelRibbonLayout.Size = new System.Drawing.Size(1302, 220);
            this.panelRibbonLayout.TabIndex = 3;
            this.panelRibbonLayout.Click += new System.EventHandler(this.ChangeFocus);
            // 
            // panelFill
            // 
            this.panelFill.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelFill.Location = new System.Drawing.Point(0, 274);
            this.panelFill.Name = "panelFill";
            this.panelFill.Size = new System.Drawing.Size(1302, 263);
            this.panelFill.TabIndex = 4;
            this.panelFill.Click += new System.EventHandler(this.ChangeFocus);
            // 
            // chooseImageDialog
            // 
            this.chooseImageDialog.FileName = "openFileDialog1";
            this.chooseImageDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.BtnImageChosen);
            // 
            // ConfiguratoreRibbon
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1302, 537);
            this.Controls.Add(this.panelFill);
            this.Controls.Add(this.panelRibbonLayout);
            this.Controls.Add(this.toolStripTopMenu);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ConfiguratoreRibbon";
            this.Text = "Form1";
            this.toolStripTopMenu.ResumeLayout(false);
            this.toolStripTopMenu.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStripTopMenu;
        private System.Windows.Forms.Panel panelRibbonLayout;
        private System.Windows.Forms.ToolStripButton AddGroup;
        private System.Windows.Forms.Panel panelFill;
        private System.Windows.Forms.ToolStripButton AddButton;
        private System.Windows.Forms.OpenFileDialog chooseImageDialog;
    }
}

