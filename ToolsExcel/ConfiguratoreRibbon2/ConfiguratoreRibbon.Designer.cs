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
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.AddGroup = new System.Windows.Forms.ToolStripButton();
            this.toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.nuovoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.scegliToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.addDropdown = new System.Windows.Forms.ToolStripButton();
            this.addEmptyContainer = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.ctrlLeftButton = new System.Windows.Forms.ToolStripButton();
            this.ctrlDownButton = new System.Windows.Forms.ToolStripButton();
            this.ctrlUpButton = new System.Windows.Forms.ToolStripButton();
            this.ctrlRightButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.panelRibbonLayout = new System.Windows.Forms.Panel();
            this.imageListSmall = new System.Windows.Forms.ImageList(this.components);
            this.imageListNormal = new System.Windows.Forms.ImageList(this.components);
            this.panelApplicazione = new System.Windows.Forms.Panel();
            this.cmbApplicazioni = new System.Windows.Forms.ComboBox();
            this.lbTitoloApplicazione = new System.Windows.Forms.Label();
            this.tableLayoutForm = new System.Windows.Forms.TableLayoutPanel();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnSalva = new System.Windows.Forms.Button();
            this.toolStripTopMenu.SuspendLayout();
            this.panelApplicazione.SuspendLayout();
            this.tableLayoutForm.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripTopMenu
            // 
            this.toolStripTopMenu.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.toolStripTopMenu.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripTopMenu.Dock = System.Windows.Forms.DockStyle.None;
            this.toolStripTopMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStripTopMenu.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.toolStripTopMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripSeparator2,
            this.AddGroup,
            this.toolStripDropDownButton1,
            this.addDropdown,
            this.addEmptyContainer,
            this.toolStripSeparator1,
            this.ctrlLeftButton,
            this.ctrlDownButton,
            this.ctrlUpButton,
            this.ctrlRightButton,
            this.toolStripSeparator3});
            this.toolStripTopMenu.Location = new System.Drawing.Point(244, 0);
            this.toolStripTopMenu.Name = "toolStripTopMenu";
            this.toolStripTopMenu.Size = new System.Drawing.Size(382, 56);
            this.toolStripTopMenu.TabIndex = 2;
            this.toolStripTopMenu.Text = "Drop down";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 56);
            // 
            // AddGroup
            // 
            this.AddGroup.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addGroup;
            this.AddGroup.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.AddGroup.Name = "AddGroup";
            this.AddGroup.Size = new System.Drawing.Size(51, 53);
            this.AddGroup.Text = "Gruppo";
            this.AddGroup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.AddGroup.Click += new System.EventHandler(this.AggiungiGruppo_Click);
            // 
            // toolStripDropDownButton1
            // 
            this.toolStripDropDownButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nuovoToolStripMenuItem,
            this.scegliToolStripMenuItem});
            this.toolStripDropDownButton1.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addButton;
            this.toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.Size = new System.Drawing.Size(49, 53);
            this.toolStripDropDownButton1.Text = "Tasto";
            this.toolStripDropDownButton1.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // nuovoToolStripMenuItem
            // 
            this.nuovoToolStripMenuItem.Name = "nuovoToolStripMenuItem";
            this.nuovoToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.nuovoToolStripMenuItem.Text = "Nuovo";
            this.nuovoToolStripMenuItem.Click += new System.EventHandler(this.AggiungiTasto_Click);
            // 
            // scegliToolStripMenuItem
            // 
            this.scegliToolStripMenuItem.Name = "scegliToolStripMenuItem";
            this.scegliToolStripMenuItem.Size = new System.Drawing.Size(168, 22);
            this.scegliToolStripMenuItem.Text = "Scegli tra esistenti";
            this.scegliToolStripMenuItem.Click += new System.EventHandler(this.scegliToolStripMenuItem_Click);
            // 
            // addDropdown
            // 
            this.addDropdown.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addDropDown;
            this.addDropdown.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.addDropdown.Name = "addDropdown";
            this.addDropdown.Size = new System.Drawing.Size(70, 53);
            this.addDropdown.Text = "Drop down";
            this.addDropdown.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.addDropdown.ToolTipText = "Aggiungi drop down";
            this.addDropdown.Click += new System.EventHandler(this.AggiungiDropDown_Click);
            // 
            // addEmptyContainer
            // 
            this.addEmptyContainer.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addEmptySlot;
            this.addEmptyContainer.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.addEmptyContainer.Name = "addEmptyContainer";
            this.addEmptyContainer.Size = new System.Drawing.Size(74, 53);
            this.addEmptyContainer.Text = "Contenitore";
            this.addEmptyContainer.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.addEmptyContainer.ToolTipText = "Aggiungi contenitore vuoto";
            this.addEmptyContainer.Click += new System.EventHandler(this.AggiungiContenitoreVuoto_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 56);
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
            this.ctrlLeftButton.Click += new System.EventHandler(this.MoveLeft_Click);
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
            this.ctrlDownButton.Click += new System.EventHandler(this.MoveDown_Click);
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
            this.ctrlUpButton.Click += new System.EventHandler(this.MoveUp_Click);
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
            this.ctrlRightButton.Click += new System.EventHandler(this.MoveRight_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 56);
            // 
            // panelRibbonLayout
            // 
            this.panelRibbonLayout.BackColor = System.Drawing.SystemColors.ControlLight;
            this.tableLayoutForm.SetColumnSpan(this.panelRibbonLayout, 100);
            this.panelRibbonLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelRibbonLayout.Location = new System.Drawing.Point(2, 58);
            this.panelRibbonLayout.Margin = new System.Windows.Forms.Padding(2);
            this.panelRibbonLayout.Name = "panelRibbonLayout";
            this.panelRibbonLayout.Padding = new System.Windows.Forms.Padding(2);
            this.panelRibbonLayout.Size = new System.Drawing.Size(1517, 208);
            this.panelRibbonLayout.TabIndex = 3;
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
            // panelApplicazione
            // 
            this.panelApplicazione.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelApplicazione.Controls.Add(this.cmbApplicazioni);
            this.panelApplicazione.Controls.Add(this.lbTitoloApplicazione);
            this.panelApplicazione.Location = new System.Drawing.Point(3, 3);
            this.panelApplicazione.Name = "panelApplicazione";
            this.panelApplicazione.Size = new System.Drawing.Size(238, 50);
            this.panelApplicazione.TabIndex = 15;
            // 
            // cmbApplicazioni
            // 
            this.cmbApplicazioni.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbApplicazioni.FormattingEnabled = true;
            this.cmbApplicazioni.Location = new System.Drawing.Point(6, 23);
            this.cmbApplicazioni.Name = "cmbApplicazioni";
            this.cmbApplicazioni.Size = new System.Drawing.Size(229, 24);
            this.cmbApplicazioni.TabIndex = 14;
            // 
            // lbTitoloApplicazione
            // 
            this.lbTitoloApplicazione.AutoSize = true;
            this.lbTitoloApplicazione.Location = new System.Drawing.Point(3, 4);
            this.lbTitoloApplicazione.Name = "lbTitoloApplicazione";
            this.lbTitoloApplicazione.Size = new System.Drawing.Size(86, 16);
            this.lbTitoloApplicazione.TabIndex = 13;
            this.lbTitoloApplicazione.Text = "Applicazione";
            // 
            // tableLayoutForm
            // 
            this.tableLayoutForm.AutoSize = true;
            this.tableLayoutForm.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutForm.ColumnCount = 5;
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 244F));
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 382F));
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 232F));
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 247F));
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 416F));
            this.tableLayoutForm.Controls.Add(this.panelRibbonLayout, 0, 1);
            this.tableLayoutForm.Controls.Add(this.toolStripTopMenu, 1, 0);
            this.tableLayoutForm.Controls.Add(this.panelBottom, 4, 2);
            this.tableLayoutForm.Controls.Add(this.panelApplicazione, 0, 0);
            this.tableLayoutForm.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutForm.MinimumSize = new System.Drawing.Size(1520, 316);
            this.tableLayoutForm.Name = "tableLayoutForm";
            this.tableLayoutForm.RowCount = 3;
            this.tableLayoutForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 56F));
            this.tableLayoutForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 212F));
            this.tableLayoutForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
            this.tableLayoutForm.Size = new System.Drawing.Size(1521, 316);
            this.tableLayoutForm.TabIndex = 16;
            // 
            // panelBottom
            // 
            this.panelBottom.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelBottom.Controls.Add(this.btnSalva);
            this.panelBottom.Location = new System.Drawing.Point(1105, 268);
            this.panelBottom.Margin = new System.Windows.Forms.Padding(0);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(416, 48);
            this.panelBottom.TabIndex = 14;
            // 
            // btnSalva
            // 
            this.btnSalva.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnSalva.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSalva.Location = new System.Drawing.Point(303, 0);
            this.btnSalva.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSalva.Name = "btnSalva";
            this.btnSalva.Size = new System.Drawing.Size(113, 48);
            this.btnSalva.TabIndex = 8;
            this.btnSalva.Text = "Salva";
            this.btnSalva.UseVisualStyleBackColor = true;
            this.btnSalva.Click += new System.EventHandler(this.btnSalva_Click);
            // 
            // ConfiguratoreRibbon
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1550, 383);
            this.Controls.Add(this.tableLayoutForm);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ConfiguratoreRibbon";
            this.Text = "Form1";
            this.toolStripTopMenu.ResumeLayout(false);
            this.toolStripTopMenu.PerformLayout();
            this.panelApplicazione.ResumeLayout(false);
            this.panelApplicazione.PerformLayout();
            this.tableLayoutForm.ResumeLayout(false);
            this.tableLayoutForm.PerformLayout();
            this.panelBottom.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStripTopMenu;
        private System.Windows.Forms.Panel panelRibbonLayout;
        private System.Windows.Forms.ToolStripButton AddGroup;
        private System.Windows.Forms.ImageList imageListSmall;
        private System.Windows.Forms.ImageList imageListNormal;
        private System.Windows.Forms.ToolStripButton ctrlDownButton;
        private System.Windows.Forms.ToolStripButton ctrlLeftButton;
        private System.Windows.Forms.ToolStripButton ctrlRightButton;
        private System.Windows.Forms.ToolStripButton ctrlUpButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton addDropdown;
        private System.Windows.Forms.ToolStripButton addEmptyContainer;
        private System.Windows.Forms.Label lbTitoloApplicazione;
        private System.Windows.Forms.ComboBox cmbApplicazioni;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.Panel panelApplicazione;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton1;
        private System.Windows.Forms.ToolStripMenuItem nuovoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem scegliToolStripMenuItem;
        private System.Windows.Forms.TableLayoutPanel tableLayoutForm;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnSalva;
    }
}

