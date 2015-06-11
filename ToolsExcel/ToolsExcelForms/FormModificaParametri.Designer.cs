namespace Iren.ToolsExcel.Forms
{
    partial class FormModificaParametri
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormModificaParametri));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.cmbEntita = new System.Windows.Forms.ComboBox();
            this.tabParametri = new System.Windows.Forms.TabControl();
            this.tabPageParD = new System.Windows.Forms.TabPage();
            this.dataGridParametriD = new System.Windows.Forms.DataGridView();
            this.contextMenuParD = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.modificareValoreContextMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.cancellaParametroContextMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.inserisciSopraContextMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.inserisciSottoContextMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.menuParD = new System.Windows.Forms.ToolStrip();
            this.elimiaTopMenu = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.inserisciSopraTopMenu = new System.Windows.Forms.ToolStripButton();
            this.inserisciSottoTopMenu = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.modificaTopMenu = new System.Windows.Forms.ToolStripButton();
            this.panelTopParD = new System.Windows.Forms.Panel();
            this.labelSelParD = new System.Windows.Forms.Label();
            this.cmbParametriD = new System.Windows.Forms.ComboBox();
            this.tabPageParH = new System.Windows.Forms.TabPage();
            this.dataGridParametriH = new System.Windows.Forms.DataGridView();
            this.contextMenuParH = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.modificaValoreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cancellaParametroToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.aggiungiNuovoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panelTopParH = new System.Windows.Forms.Panel();
            this.labelSelParH = new System.Windows.Forms.Label();
            this.cmbParametriH = new System.Windows.Forms.ComboBox();
            this.panelTop = new System.Windows.Forms.Panel();
            this.labelSelEntita = new System.Windows.Forms.Label();
            this.panelTopMenu = new System.Windows.Forms.Panel();
            this.tabParametri.SuspendLayout();
            this.tabPageParD.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriD)).BeginInit();
            this.contextMenuParD.SuspendLayout();
            this.menuParD.SuspendLayout();
            this.panelTopParD.SuspendLayout();
            this.tabPageParH.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriH)).BeginInit();
            this.contextMenuParH.SuspendLayout();
            this.panelTopParH.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.panelTopMenu.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbEntita
            // 
            this.cmbEntita.FormattingEnabled = true;
            this.cmbEntita.Location = new System.Drawing.Point(12, 34);
            this.cmbEntita.Name = "cmbEntita";
            this.cmbEntita.Size = new System.Drawing.Size(486, 28);
            this.cmbEntita.TabIndex = 0;
            this.cmbEntita.SelectedIndexChanged += new System.EventHandler(this.cmbEntita_SelectedIndexChanged);
            // 
            // tabParametri
            // 
            this.tabParametri.Controls.Add(this.tabPageParD);
            this.tabParametri.Controls.Add(this.tabPageParH);
            this.tabParametri.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabParametri.Location = new System.Drawing.Point(2, 75);
            this.tabParametri.Margin = new System.Windows.Forms.Padding(0);
            this.tabParametri.Name = "tabParametri";
            this.tabParametri.SelectedIndex = 0;
            this.tabParametri.Size = new System.Drawing.Size(776, 428);
            this.tabParametri.TabIndex = 1;
            // 
            // tabPageParD
            // 
            this.tabPageParD.Controls.Add(this.dataGridParametriD);
            this.tabPageParD.Controls.Add(this.panelTopMenu);
            this.tabPageParD.Controls.Add(this.panelTopParD);
            this.tabPageParD.Location = new System.Drawing.Point(4, 29);
            this.tabPageParD.Margin = new System.Windows.Forms.Padding(0);
            this.tabPageParD.Name = "tabPageParD";
            this.tabPageParD.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageParD.Size = new System.Drawing.Size(768, 395);
            this.tabPageParD.TabIndex = 0;
            this.tabPageParD.Text = "Parametri Giornalieri";
            this.tabPageParD.UseVisualStyleBackColor = true;
            // 
            // dataGridParametriD
            // 
            this.dataGridParametriD.AllowUserToAddRows = false;
            this.dataGridParametriD.AllowUserToDeleteRows = false;
            this.dataGridParametriD.AllowUserToResizeColumns = false;
            this.dataGridParametriD.AllowUserToResizeRows = false;
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(216)))), ((int)(((byte)(251)))), ((int)(((byte)(252)))));
            this.dataGridParametriD.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridParametriD.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridParametriD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridParametriD.ContextMenuStrip = this.contextMenuParD;
            this.dataGridParametriD.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridParametriD.Location = new System.Drawing.Point(3, 84);
            this.dataGridParametriD.MultiSelect = false;
            this.dataGridParametriD.Name = "dataGridParametriD";
            this.dataGridParametriD.Size = new System.Drawing.Size(762, 308);
            this.dataGridParametriD.TabIndex = 6;
            this.dataGridParametriD.VirtualMode = true;
            this.dataGridParametriD.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridParametriD_CellBeginEdit);
            this.dataGridParametriD.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridParametriD_CellValidating);
            this.dataGridParametriD.CurrentCellChanged += new System.EventHandler(this.dataGridParametriD_CurrentCellChanged);
            this.dataGridParametriD.RowDirtyStateNeeded += new System.Windows.Forms.QuestionEventHandler(this.dataGridParametriD_RowDirtyStateNeeded);
            this.dataGridParametriD.RowValidated += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridParametriD_RowValidated);
            this.dataGridParametriD.RowValidating += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridParametriD_RowValidating);
            this.dataGridParametriD.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridParametriD_MouseDown);
            this.dataGridParametriD.MouseEnter += new System.EventHandler(this.dataGridParametriD_MouseEnter);
            // 
            // contextMenuParD
            // 
            this.contextMenuParD.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.modificareValoreContextMenu,
            this.cancellaParametroContextMenu,
            this.inserisciSopraContextMenu,
            this.inserisciSottoContextMenu});
            this.contextMenuParD.Name = "contextMenuDataGrid";
            this.contextMenuParD.Size = new System.Drawing.Size(233, 92);
            // 
            // modificareValoreContextMenu
            // 
            this.modificareValoreContextMenu.Image = ((System.Drawing.Image)(resources.GetObject("modificareValoreContextMenu.Image")));
            this.modificareValoreContextMenu.Name = "modificareValoreContextMenu";
            this.modificareValoreContextMenu.Size = new System.Drawing.Size(232, 22);
            this.modificareValoreContextMenu.Text = "Modificare valore";
            this.modificareValoreContextMenu.ToolTipText = "Modifica il valore del parametro se il parametro non è ancora attivo";
            this.modificareValoreContextMenu.Click += new System.EventHandler(this.modificareValoreContextMenu_Click);
            // 
            // cancellaParametroContextMenu
            // 
            this.cancellaParametroContextMenu.Image = ((System.Drawing.Image)(resources.GetObject("cancellaParametroContextMenu.Image")));
            this.cancellaParametroContextMenu.Name = "cancellaParametroContextMenu";
            this.cancellaParametroContextMenu.Size = new System.Drawing.Size(232, 22);
            this.cancellaParametroContextMenu.Text = "Elimina parametro";
            this.cancellaParametroContextMenu.ToolTipText = "Cancella il parametro se il parametro non è ancora attivo";
            this.cancellaParametroContextMenu.Click += new System.EventHandler(this.cancellaParametroContextMenu_Click);
            // 
            // inserisciSopraContextMenu
            // 
            this.inserisciSopraContextMenu.Image = ((System.Drawing.Image)(resources.GetObject("inserisciSopraContextMenu.Image")));
            this.inserisciSopraContextMenu.Name = "inserisciSopraContextMenu";
            this.inserisciSopraContextMenu.Size = new System.Drawing.Size(232, 22);
            this.inserisciSopraContextMenu.Text = "Inserisci sopra riga selezionata";
            this.inserisciSopraContextMenu.Click += new System.EventHandler(this.inserisciSopraContextMenu_Click);
            // 
            // inserisciSottoContextMenu
            // 
            this.inserisciSottoContextMenu.Image = ((System.Drawing.Image)(resources.GetObject("inserisciSottoContextMenu.Image")));
            this.inserisciSottoContextMenu.Name = "inserisciSottoContextMenu";
            this.inserisciSottoContextMenu.Size = new System.Drawing.Size(232, 22);
            this.inserisciSottoContextMenu.Text = "Inserisci sotto riga selezionata";
            this.inserisciSottoContextMenu.Click += new System.EventHandler(this.inserisciSottoContextMenu_Click);
            // 
            // menuParD
            // 
            this.menuParD.BackColor = System.Drawing.SystemColors.Control;
            this.menuParD.Dock = System.Windows.Forms.DockStyle.Fill;
            this.menuParD.GripMargin = new System.Windows.Forms.Padding(3);
            this.menuParD.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.menuParD.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.elimiaTopMenu,
            this.toolStripSeparator1,
            this.inserisciSopraTopMenu,
            this.inserisciSottoTopMenu,
            this.toolStripSeparator2,
            this.modificaTopMenu});
            this.menuParD.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow;
            this.menuParD.Location = new System.Drawing.Point(0, 0);
            this.menuParD.Name = "menuParD";
            this.menuParD.Padding = new System.Windows.Forms.Padding(3);
            this.menuParD.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.menuParD.Size = new System.Drawing.Size(762, 37);
            this.menuParD.TabIndex = 10;
            this.menuParD.Text = "Strumenti Parametri Giornalieri";
            // 
            // elimiaTopMenu
            // 
            this.elimiaTopMenu.Image = ((System.Drawing.Image)(resources.GetObject("elimiaTopMenu.Image")));
            this.elimiaTopMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.elimiaTopMenu.Name = "elimiaTopMenu";
            this.elimiaTopMenu.Size = new System.Drawing.Size(66, 28);
            this.elimiaTopMenu.Text = "Elimina";
            this.elimiaTopMenu.ToolTipText = "Elimina la riga selezionata";
            this.elimiaTopMenu.Click += new System.EventHandler(this.elimiaTopMenu_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 31);
            // 
            // inserisciSopraTopMenu
            // 
            this.inserisciSopraTopMenu.Image = ((System.Drawing.Image)(resources.GetObject("inserisciSopraTopMenu.Image")));
            this.inserisciSopraTopMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.inserisciSopraTopMenu.Name = "inserisciSopraTopMenu";
            this.inserisciSopraTopMenu.Size = new System.Drawing.Size(101, 28);
            this.inserisciSopraTopMenu.Text = "Inserisci sopra";
            this.inserisciSopraTopMenu.ToolTipText = "Inserisci sopra la riga corrente";
            this.inserisciSopraTopMenu.Click += new System.EventHandler(this.inserisciSopraTopMenu_Click);
            // 
            // inserisciSottoTopMenu
            // 
            this.inserisciSottoTopMenu.Image = ((System.Drawing.Image)(resources.GetObject("inserisciSottoTopMenu.Image")));
            this.inserisciSottoTopMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.inserisciSottoTopMenu.Name = "inserisciSottoTopMenu";
            this.inserisciSottoTopMenu.Size = new System.Drawing.Size(99, 28);
            this.inserisciSottoTopMenu.Text = "Inserisci sotto";
            this.inserisciSottoTopMenu.ToolTipText = "Inserisci sotto la riga corrente";
            this.inserisciSottoTopMenu.Click += new System.EventHandler(this.inserisciSottoTopMenu_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 31);
            // 
            // modificaTopMenu
            // 
            this.modificaTopMenu.Image = ((System.Drawing.Image)(resources.GetObject("modificaTopMenu.Image")));
            this.modificaTopMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.modificaTopMenu.Name = "modificaTopMenu";
            this.modificaTopMenu.Size = new System.Drawing.Size(74, 28);
            this.modificaTopMenu.Text = "Modifica";
            this.modificaTopMenu.ToolTipText = "Modifica la riga selezionata";
            this.modificaTopMenu.Click += new System.EventHandler(this.modificaTopMenu_Click);
            // 
            // panelTopParD
            // 
            this.panelTopParD.Controls.Add(this.labelSelParD);
            this.panelTopParD.Controls.Add(this.cmbParametriD);
            this.panelTopParD.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTopParD.Location = new System.Drawing.Point(3, 3);
            this.panelTopParD.Name = "panelTopParD";
            this.panelTopParD.Size = new System.Drawing.Size(762, 44);
            this.panelTopParD.TabIndex = 9;
            // 
            // labelSelParD
            // 
            this.labelSelParD.AutoSize = true;
            this.labelSelParD.Location = new System.Drawing.Point(5, 11);
            this.labelSelParD.Name = "labelSelParD";
            this.labelSelParD.Size = new System.Drawing.Size(170, 20);
            this.labelSelParD.TabIndex = 6;
            this.labelSelParD.Text = "Seleziona il parametro:";
            // 
            // cmbParametriD
            // 
            this.cmbParametriD.FormattingEnabled = true;
            this.cmbParametriD.Location = new System.Drawing.Point(181, 8);
            this.cmbParametriD.Name = "cmbParametriD";
            this.cmbParametriD.Size = new System.Drawing.Size(310, 28);
            this.cmbParametriD.TabIndex = 5;
            this.cmbParametriD.SelectedIndexChanged += new System.EventHandler(this.cmbParametriD_SelectedIndexChanged);
            // 
            // tabPageParH
            // 
            this.tabPageParH.Controls.Add(this.dataGridParametriH);
            this.tabPageParH.Controls.Add(this.panelTopParH);
            this.tabPageParH.Location = new System.Drawing.Point(4, 29);
            this.tabPageParH.Name = "tabPageParH";
            this.tabPageParH.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageParH.Size = new System.Drawing.Size(768, 395);
            this.tabPageParH.TabIndex = 1;
            this.tabPageParH.Text = "Parametri Orari";
            this.tabPageParH.UseVisualStyleBackColor = true;
            // 
            // dataGridParametriH
            // 
            this.dataGridParametriH.AllowUserToAddRows = false;
            this.dataGridParametriH.AllowUserToDeleteRows = false;
            this.dataGridParametriH.AllowUserToResizeColumns = false;
            this.dataGridParametriH.AllowUserToResizeRows = false;
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(216)))), ((int)(((byte)(251)))), ((int)(((byte)(252)))));
            this.dataGridParametriH.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dataGridParametriH.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridParametriH.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridParametriH.ContextMenuStrip = this.contextMenuParH;
            this.dataGridParametriH.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridParametriH.Location = new System.Drawing.Point(3, 50);
            this.dataGridParametriH.MultiSelect = false;
            this.dataGridParametriH.Name = "dataGridParametriH";
            this.dataGridParametriH.Size = new System.Drawing.Size(762, 342);
            this.dataGridParametriH.TabIndex = 6;
            // 
            // contextMenuParH
            // 
            this.contextMenuParH.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.modificaValoreToolStripMenuItem,
            this.cancellaParametroToolStripMenuItem1,
            this.aggiungiNuovoToolStripMenuItem});
            this.contextMenuParH.Name = "contextMenuParH";
            this.contextMenuParH.Size = new System.Drawing.Size(178, 70);
            // 
            // modificaValoreToolStripMenuItem
            // 
            this.modificaValoreToolStripMenuItem.Name = "modificaValoreToolStripMenuItem";
            this.modificaValoreToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.modificaValoreToolStripMenuItem.Text = "Modifica valore";
            // 
            // cancellaParametroToolStripMenuItem1
            // 
            this.cancellaParametroToolStripMenuItem1.Name = "cancellaParametroToolStripMenuItem1";
            this.cancellaParametroToolStripMenuItem1.Size = new System.Drawing.Size(177, 22);
            this.cancellaParametroToolStripMenuItem1.Text = "Cancella parametro";
            // 
            // aggiungiNuovoToolStripMenuItem
            // 
            this.aggiungiNuovoToolStripMenuItem.Name = "aggiungiNuovoToolStripMenuItem";
            this.aggiungiNuovoToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.aggiungiNuovoToolStripMenuItem.Text = "Aggiungi nuovo";
            // 
            // panelTopParH
            // 
            this.panelTopParH.Controls.Add(this.labelSelParH);
            this.panelTopParH.Controls.Add(this.cmbParametriH);
            this.panelTopParH.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTopParH.Location = new System.Drawing.Point(3, 3);
            this.panelTopParH.Name = "panelTopParH";
            this.panelTopParH.Size = new System.Drawing.Size(762, 47);
            this.panelTopParH.TabIndex = 8;
            // 
            // labelSelParH
            // 
            this.labelSelParH.AutoSize = true;
            this.labelSelParH.Location = new System.Drawing.Point(5, 11);
            this.labelSelParH.Name = "labelSelParH";
            this.labelSelParH.Size = new System.Drawing.Size(170, 20);
            this.labelSelParH.TabIndex = 7;
            this.labelSelParH.Text = "Seleziona il parametro:";
            // 
            // cmbParametriH
            // 
            this.cmbParametriH.FormattingEnabled = true;
            this.cmbParametriH.Location = new System.Drawing.Point(181, 8);
            this.cmbParametriH.Name = "cmbParametriH";
            this.cmbParametriH.Size = new System.Drawing.Size(310, 28);
            this.cmbParametriH.TabIndex = 5;
            this.cmbParametriH.SelectedIndexChanged += new System.EventHandler(this.cmbParametriH_SelectedIndexChanged);
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.labelSelEntita);
            this.panelTop.Controls.Add(this.cmbEntita);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(2, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(776, 75);
            this.panelTop.TabIndex = 2;
            // 
            // labelSelEntita
            // 
            this.labelSelEntita.AutoSize = true;
            this.labelSelEntita.Location = new System.Drawing.Point(12, 9);
            this.labelSelEntita.Name = "labelSelEntita";
            this.labelSelEntita.Size = new System.Drawing.Size(115, 20);
            this.labelSelEntita.TabIndex = 1;
            this.labelSelEntita.Text = "Seleziona l\'UP:";
            // 
            // panelTopMenu
            // 
            this.panelTopMenu.Controls.Add(this.menuParD);
            this.panelTopMenu.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTopMenu.Location = new System.Drawing.Point(3, 47);
            this.panelTopMenu.Name = "panelTopMenu";
            this.panelTopMenu.Size = new System.Drawing.Size(762, 37);
            this.panelTopMenu.TabIndex = 11;
            // 
            // FormModificaParametri
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(778, 504);
            this.Controls.Add(this.tabParametri);
            this.Controls.Add(this.panelTop);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormModificaParametri";
            this.Padding = new System.Windows.Forms.Padding(2, 0, 0, 1);
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.tabParametri.ResumeLayout(false);
            this.tabPageParD.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriD)).EndInit();
            this.contextMenuParD.ResumeLayout(false);
            this.menuParD.ResumeLayout(false);
            this.menuParD.PerformLayout();
            this.panelTopParD.ResumeLayout(false);
            this.panelTopParD.PerformLayout();
            this.tabPageParH.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriH)).EndInit();
            this.contextMenuParH.ResumeLayout(false);
            this.panelTopParH.ResumeLayout(false);
            this.panelTopParH.PerformLayout();
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.panelTopMenu.ResumeLayout(false);
            this.panelTopMenu.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbEntita;
        private System.Windows.Forms.TabControl tabParametri;
        private System.Windows.Forms.TabPage tabPageParD;
        private System.Windows.Forms.DataGridView dataGridParametriD;
        private System.Windows.Forms.ComboBox cmbParametriD;
        private System.Windows.Forms.TabPage tabPageParH;
        private System.Windows.Forms.DataGridView dataGridParametriH;
        private System.Windows.Forms.ComboBox cmbParametriH;
        private System.Windows.Forms.Panel panelTopParD;
        private System.Windows.Forms.Panel panelTopParH;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.ContextMenuStrip contextMenuParD;
        private System.Windows.Forms.ToolStripMenuItem modificareValoreContextMenu;
        private System.Windows.Forms.ToolStripMenuItem cancellaParametroContextMenu;
        private System.Windows.Forms.Label labelSelParD;
        private System.Windows.Forms.Label labelSelEntita;
        private System.Windows.Forms.Label labelSelParH;
        private System.Windows.Forms.ToolStripMenuItem inserisciSopraContextMenu;
        private System.Windows.Forms.ToolStripMenuItem inserisciSottoContextMenu;
        private System.Windows.Forms.ContextMenuStrip contextMenuParH;
        private System.Windows.Forms.ToolStripMenuItem modificaValoreToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cancellaParametroToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem aggiungiNuovoToolStripMenuItem;
        private System.Windows.Forms.ToolStrip menuParD;
        private System.Windows.Forms.ToolStripButton elimiaTopMenu;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton inserisciSopraTopMenu;
        private System.Windows.Forms.ToolStripButton inserisciSottoTopMenu;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton modificaTopMenu;
        private System.Windows.Forms.Panel panelTopMenu;
    }
}