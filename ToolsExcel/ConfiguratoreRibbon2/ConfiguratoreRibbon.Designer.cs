namespace Iren.ToolsExcel.ConfiguratoreRibbon
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
            this.components = new System.ComponentModel.Container();
            this.toolStripTopMenu = new System.Windows.Forms.ToolStrip();
            this.AddGroup = new System.Windows.Forms.ToolStripButton();
            this.AddButton = new System.Windows.Forms.ToolStripButton();
            this.addDropdown = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.ctrlLeftButton = new System.Windows.Forms.ToolStripButton();
            this.ctrlDownButton = new System.Windows.Forms.ToolStripButton();
            this.ctrlUpButton = new System.Windows.Forms.ToolStripButton();
            this.ctrlRightButton = new System.Windows.Forms.ToolStripButton();
            this.panelRibbonLayout = new System.Windows.Forms.Panel();
            this.panelFill = new System.Windows.Forms.Panel();
            this.imageListSmall = new System.Windows.Forms.ImageList(this.components);
            this.imageListNormal = new System.Windows.Forms.ImageList(this.components);
            this.toolStripTopMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripTopMenu
            // 
            this.toolStripTopMenu.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripTopMenu.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.toolStripTopMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.AddGroup,
            this.AddButton,
            this.addDropdown,
            this.toolStripSeparator1,
            this.ctrlLeftButton,
            this.ctrlDownButton,
            this.ctrlUpButton,
            this.ctrlRightButton});
            this.toolStripTopMenu.Location = new System.Drawing.Point(0, 0);
            this.toolStripTopMenu.Name = "toolStripTopMenu";
            this.toolStripTopMenu.Size = new System.Drawing.Size(1521, 54);
            this.toolStripTopMenu.TabIndex = 2;
            this.toolStripTopMenu.Text = "Drop down";
            // 
            // AddGroup
            // 
            this.AddGroup.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addGroup;
            this.AddGroup.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.AddGroup.Name = "AddGroup";
            this.AddGroup.Size = new System.Drawing.Size(51, 51);
            this.AddGroup.Text = "Gruppo";
            this.AddGroup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.AddGroup.Click += new System.EventHandler(this.AggiungiGruppo_Click);
            // 
            // AddButton
            // 
            this.AddButton.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addButton;
            this.AddButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.AddButton.Name = "AddButton";
            this.AddButton.Size = new System.Drawing.Size(40, 51);
            this.AddButton.Text = "Tasto";
            this.AddButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.AddButton.Click += new System.EventHandler(this.AggiungiTasto_Click);
            // 
            // addDropdown
            // 
            this.addDropdown.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addDropDown;
            this.addDropdown.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.addDropdown.Name = "addDropdown";
            this.addDropdown.Size = new System.Drawing.Size(70, 51);
            this.addDropdown.Text = "Drop down";
            this.addDropdown.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.addDropdown.ToolTipText = "Aggiungi drop down";
            this.addDropdown.Click += new System.EventHandler(this.AddDropDown_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 54);
            // 
            // ctrlLeftButton
            // 
            this.ctrlLeftButton.AutoSize = false;
            this.ctrlLeftButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ctrlLeftButton.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.leftArrow;
            this.ctrlLeftButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlLeftButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ctrlLeftButton.Name = "ctrlLeftButton";
            this.ctrlLeftButton.Size = new System.Drawing.Size(24, 24);
            this.ctrlLeftButton.Text = "toolStripButton2";
            this.ctrlLeftButton.ToolTipText = "Sposta il controllo a sinistra";
            this.ctrlLeftButton.Click += new System.EventHandler(this.ctrlLeftButton_Click);
            // 
            // ctrlDownButton
            // 
            this.ctrlDownButton.AutoSize = false;
            this.ctrlDownButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ctrlDownButton.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.downArrow;
            this.ctrlDownButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlDownButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ctrlDownButton.Name = "ctrlDownButton";
            this.ctrlDownButton.Size = new System.Drawing.Size(24, 24);
            this.ctrlDownButton.Text = "toolStripButton1";
            this.ctrlDownButton.ToolTipText = "Sposta il controllo in basso";
            this.ctrlDownButton.Click += new System.EventHandler(this.ctrlDownButton_Click);
            // 
            // ctrlUpButton
            // 
            this.ctrlUpButton.AutoSize = false;
            this.ctrlUpButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ctrlUpButton.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.upArrow;
            this.ctrlUpButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlUpButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ctrlUpButton.Name = "ctrlUpButton";
            this.ctrlUpButton.Size = new System.Drawing.Size(24, 24);
            this.ctrlUpButton.Text = "toolStripButton4";
            this.ctrlUpButton.ToolTipText = "Sposta il controllo in alto";
            this.ctrlUpButton.Click += new System.EventHandler(this.ctrlUpButton_Click);
            // 
            // ctrlRightButton
            // 
            this.ctrlRightButton.AutoSize = false;
            this.ctrlRightButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ctrlRightButton.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.rightArrow;
            this.ctrlRightButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlRightButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ctrlRightButton.Name = "ctrlRightButton";
            this.ctrlRightButton.Size = new System.Drawing.Size(24, 24);
            this.ctrlRightButton.Text = "toolStripButton3";
            this.ctrlRightButton.ToolTipText = "Sposta il controllo a destra";
            // 
            // panelRibbonLayout
            // 
            this.panelRibbonLayout.BackColor = System.Drawing.SystemColors.ControlLight;
            this.panelRibbonLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelRibbonLayout.Location = new System.Drawing.Point(0, 54);
            this.panelRibbonLayout.Margin = new System.Windows.Forms.Padding(4);
            this.panelRibbonLayout.Name = "panelRibbonLayout";
            this.panelRibbonLayout.Padding = new System.Windows.Forms.Padding(4);
            this.panelRibbonLayout.Size = new System.Drawing.Size(1521, 220);
            this.panelRibbonLayout.TabIndex = 3;
            // 
            // panelFill
            // 
            this.panelFill.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelFill.Location = new System.Drawing.Point(0, 274);
            this.panelFill.Name = "panelFill";
            this.panelFill.Size = new System.Drawing.Size(1521, 263);
            this.panelFill.TabIndex = 4;
            // 
            // imageListSmall
            // 
            this.imageListSmall.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageListSmall.ImageSize = new System.Drawing.Size(16, 16);
            this.imageListSmall.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // imageListNormal
            // 
            this.imageListNormal.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageListNormal.ImageSize = new System.Drawing.Size(32, 32);
            this.imageListNormal.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // ConfiguratoreRibbon
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1521, 537);
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
        private System.Windows.Forms.ImageList imageListSmall;
        private System.Windows.Forms.ImageList imageListNormal;
        private System.Windows.Forms.ToolStripButton ctrlDownButton;
        private System.Windows.Forms.ToolStripButton ctrlLeftButton;
        private System.Windows.Forms.ToolStripButton ctrlRightButton;
        private System.Windows.Forms.ToolStripButton ctrlUpButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton addDropdown;
    }
}

