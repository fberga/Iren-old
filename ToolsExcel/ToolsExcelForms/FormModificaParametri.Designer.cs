﻿namespace Iren.ToolsExcel.Forms
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.cmbEntita = new System.Windows.Forms.ComboBox();
            this.tabParametri = new System.Windows.Forms.TabControl();
            this.tabPageParD = new System.Windows.Forms.TabPage();
            this.dataGridParametriD = new System.Windows.Forms.DataGridView();
            this.contextMenuDataGrid = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.modificareValoreToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cancellaParametroToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aggiungiNuovoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.panelTopParD = new System.Windows.Forms.Panel();
            this.cmbParametriD = new System.Windows.Forms.ComboBox();
            this.panelParD = new System.Windows.Forms.Panel();
            this.btnAggiungiParD = new System.Windows.Forms.Button();
            this.tabPageParH = new System.Windows.Forms.TabPage();
            this.dataGridParametriH = new System.Windows.Forms.DataGridView();
            this.panelTopParH = new System.Windows.Forms.Panel();
            this.cmbParametriH = new System.Windows.Forms.ComboBox();
            this.panelParH = new System.Windows.Forms.Panel();
            this.btnAggiungiParH = new System.Windows.Forms.Button();
            this.panelTop = new System.Windows.Forms.Panel();
            this.tabParametri.SuspendLayout();
            this.tabPageParD.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriD)).BeginInit();
            this.contextMenuDataGrid.SuspendLayout();
            this.panelTopParD.SuspendLayout();
            this.panelParD.SuspendLayout();
            this.tabPageParH.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriH)).BeginInit();
            this.panelTopParH.SuspendLayout();
            this.panelParH.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbEntita
            // 
            this.cmbEntita.FormattingEnabled = true;
            this.cmbEntita.Location = new System.Drawing.Point(3, 12);
            this.cmbEntita.Name = "cmbEntita";
            this.cmbEntita.Size = new System.Drawing.Size(294, 28);
            this.cmbEntita.TabIndex = 0;
            this.cmbEntita.SelectedIndexChanged += new System.EventHandler(this.cmbEntita_SelectedIndexChanged);
            // 
            // tabParametri
            // 
            this.tabParametri.Controls.Add(this.tabPageParD);
            this.tabParametri.Controls.Add(this.tabPageParH);
            this.tabParametri.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabParametri.Location = new System.Drawing.Point(0, 49);
            this.tabParametri.Margin = new System.Windows.Forms.Padding(0);
            this.tabParametri.Name = "tabParametri";
            this.tabParametri.SelectedIndex = 0;
            this.tabParametri.Size = new System.Drawing.Size(778, 505);
            this.tabParametri.TabIndex = 1;
            // 
            // tabPageParD
            // 
            this.tabPageParD.Controls.Add(this.dataGridParametriD);
            this.tabPageParD.Controls.Add(this.panelTopParD);
            this.tabPageParD.Controls.Add(this.panelParD);
            this.tabPageParD.Location = new System.Drawing.Point(4, 29);
            this.tabPageParD.Margin = new System.Windows.Forms.Padding(0);
            this.tabPageParD.Name = "tabPageParD";
            this.tabPageParD.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageParD.Size = new System.Drawing.Size(770, 472);
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
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.dataGridParametriD.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridParametriD.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridParametriD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridParametriD.ContextMenuStrip = this.contextMenuDataGrid;
            this.dataGridParametriD.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridParametriD.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridParametriD.Location = new System.Drawing.Point(3, 50);
            this.dataGridParametriD.MultiSelect = false;
            this.dataGridParametriD.Name = "dataGridParametriD";
            this.dataGridParametriD.ReadOnly = true;
            this.dataGridParametriD.Size = new System.Drawing.Size(764, 373);
            this.dataGridParametriD.TabIndex = 6;
            this.dataGridParametriD.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridCellBeginEdit);
            this.dataGridParametriD.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridCellEndEdit);
            this.dataGridParametriD.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridParametriD_CellValidating);
            this.dataGridParametriD.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridParametriD_MouseDown);
            this.dataGridParametriD.MouseEnter += new System.EventHandler(this.DataGridMouseEnter);
            this.dataGridParametriD.MouseMove += new System.Windows.Forms.MouseEventHandler(this.dataGridParametriD_MouseMove);
            // 
            // contextMenuDataGrid
            // 
            this.contextMenuDataGrid.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.modificareValoreToolStripMenuItem,
            this.cancellaParametroToolStripMenuItem,
            this.aggiungiNuovoToolStripMenuItem});
            this.contextMenuDataGrid.Name = "contextMenuDataGrid";
            this.contextMenuDataGrid.ShowImageMargin = false;
            this.contextMenuDataGrid.Size = new System.Drawing.Size(153, 70);
            // 
            // modificareValoreToolStripMenuItem
            // 
            this.modificareValoreToolStripMenuItem.Name = "modificareValoreToolStripMenuItem";
            this.modificareValoreToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.modificareValoreToolStripMenuItem.Text = "Modificare valore";
            this.modificareValoreToolStripMenuItem.ToolTipText = "Modifica il valore del parametro se il parametro non è ancora attivo";
            this.modificareValoreToolStripMenuItem.Click += new System.EventHandler(this.modificareValoreToolStripMenuItem_Click);
            // 
            // cancellaParametroToolStripMenuItem
            // 
            this.cancellaParametroToolStripMenuItem.Name = "cancellaParametroToolStripMenuItem";
            this.cancellaParametroToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.cancellaParametroToolStripMenuItem.Text = "Cancella parametro";
            this.cancellaParametroToolStripMenuItem.ToolTipText = "Cancella il parametro se il parametro non è ancora attivo";
            // 
            // aggiungiNuovoToolStripMenuItem
            // 
            this.aggiungiNuovoToolStripMenuItem.Name = "aggiungiNuovoToolStripMenuItem";
            this.aggiungiNuovoToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.aggiungiNuovoToolStripMenuItem.Text = "Aggiungi nuovo";
            this.aggiungiNuovoToolStripMenuItem.ToolTipText = "Aggiunge il nuovo valore con validità dal giorno successivo all\'inserimento";
            // 
            // panelTopParD
            // 
            this.panelTopParD.Controls.Add(this.cmbParametriD);
            this.panelTopParD.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTopParD.Location = new System.Drawing.Point(3, 3);
            this.panelTopParD.Name = "panelTopParD";
            this.panelTopParD.Size = new System.Drawing.Size(764, 47);
            this.panelTopParD.TabIndex = 9;
            // 
            // cmbParametriD
            // 
            this.cmbParametriD.FormattingEnabled = true;
            this.cmbParametriD.Location = new System.Drawing.Point(3, 3);
            this.cmbParametriD.Name = "cmbParametriD";
            this.cmbParametriD.Size = new System.Drawing.Size(294, 28);
            this.cmbParametriD.TabIndex = 5;
            this.cmbParametriD.SelectedIndexChanged += new System.EventHandler(this.cmbParametriD_SelectedIndexChanged);
            // 
            // panelParD
            // 
            this.panelParD.Controls.Add(this.btnAggiungiParD);
            this.panelParD.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelParD.Location = new System.Drawing.Point(3, 423);
            this.panelParD.Name = "panelParD";
            this.panelParD.Padding = new System.Windows.Forms.Padding(0, 0, 3, 0);
            this.panelParD.Size = new System.Drawing.Size(764, 46);
            this.panelParD.TabIndex = 7;
            // 
            // btnAggiungiParD
            // 
            this.btnAggiungiParD.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAggiungiParD.Location = new System.Drawing.Point(715, 0);
            this.btnAggiungiParD.Name = "btnAggiungiParD";
            this.btnAggiungiParD.Size = new System.Drawing.Size(46, 46);
            this.btnAggiungiParD.TabIndex = 0;
            this.btnAggiungiParD.Text = "+";
            this.btnAggiungiParD.UseVisualStyleBackColor = true;
            // 
            // tabPageParH
            // 
            this.tabPageParH.Controls.Add(this.dataGridParametriH);
            this.tabPageParH.Controls.Add(this.panelTopParH);
            this.tabPageParH.Controls.Add(this.panelParH);
            this.tabPageParH.Location = new System.Drawing.Point(4, 29);
            this.tabPageParH.Name = "tabPageParH";
            this.tabPageParH.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageParH.Size = new System.Drawing.Size(770, 472);
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
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.dataGridParametriH.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridParametriH.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridParametriH.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridParametriH.ContextMenuStrip = this.contextMenuDataGrid;
            this.dataGridParametriH.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridParametriH.Location = new System.Drawing.Point(3, 50);
            this.dataGridParametriH.MultiSelect = false;
            this.dataGridParametriH.Name = "dataGridParametriH";
            this.dataGridParametriH.Size = new System.Drawing.Size(764, 373);
            this.dataGridParametriH.TabIndex = 6;
            this.dataGridParametriH.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridCellBeginEdit);
            this.dataGridParametriH.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridCellEndEdit);
            this.dataGridParametriH.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridParametriH_CellValidating);
            this.dataGridParametriH.MouseDown += new System.Windows.Forms.MouseEventHandler(this.dataGridParametriH_MouseDown);
            this.dataGridParametriH.MouseEnter += new System.EventHandler(this.DataGridMouseEnter);
            this.dataGridParametriH.MouseMove += new System.Windows.Forms.MouseEventHandler(this.dataGridParametriH_MouseMove);
            // 
            // panelTopParH
            // 
            this.panelTopParH.Controls.Add(this.cmbParametriH);
            this.panelTopParH.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTopParH.Location = new System.Drawing.Point(3, 3);
            this.panelTopParH.Name = "panelTopParH";
            this.panelTopParH.Size = new System.Drawing.Size(764, 47);
            this.panelTopParH.TabIndex = 8;
            // 
            // cmbParametriH
            // 
            this.cmbParametriH.FormattingEnabled = true;
            this.cmbParametriH.Location = new System.Drawing.Point(3, 3);
            this.cmbParametriH.Name = "cmbParametriH";
            this.cmbParametriH.Size = new System.Drawing.Size(294, 28);
            this.cmbParametriH.TabIndex = 5;
            this.cmbParametriH.SelectedIndexChanged += new System.EventHandler(this.cmbParametriH_SelectedIndexChanged);
            // 
            // panelParH
            // 
            this.panelParH.Controls.Add(this.btnAggiungiParH);
            this.panelParH.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelParH.Location = new System.Drawing.Point(3, 423);
            this.panelParH.Name = "panelParH";
            this.panelParH.Padding = new System.Windows.Forms.Padding(0, 0, 3, 0);
            this.panelParH.Size = new System.Drawing.Size(764, 46);
            this.panelParH.TabIndex = 7;
            // 
            // btnAggiungiParH
            // 
            this.btnAggiungiParH.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAggiungiParH.Location = new System.Drawing.Point(715, 0);
            this.btnAggiungiParH.Name = "btnAggiungiParH";
            this.btnAggiungiParH.Size = new System.Drawing.Size(46, 46);
            this.btnAggiungiParH.TabIndex = 0;
            this.btnAggiungiParH.Text = "+";
            this.btnAggiungiParH.UseVisualStyleBackColor = true;
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.cmbEntita);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(778, 49);
            this.panelTop.TabIndex = 2;
            // 
            // FormModificaParametri
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(778, 554);
            this.Controls.Add(this.tabParametri);
            this.Controls.Add(this.panelTop);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormModificaParametri";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.tabParametri.ResumeLayout(false);
            this.tabPageParD.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriD)).EndInit();
            this.contextMenuDataGrid.ResumeLayout(false);
            this.panelTopParD.ResumeLayout(false);
            this.panelParD.ResumeLayout(false);
            this.tabPageParH.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriH)).EndInit();
            this.panelTopParH.ResumeLayout(false);
            this.panelParH.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbEntita;
        private System.Windows.Forms.TabControl tabParametri;
        private System.Windows.Forms.TabPage tabPageParD;
        private System.Windows.Forms.Panel panelParD;
        private System.Windows.Forms.Button btnAggiungiParD;
        private System.Windows.Forms.DataGridView dataGridParametriD;
        private System.Windows.Forms.ComboBox cmbParametriD;
        private System.Windows.Forms.TabPage tabPageParH;
        private System.Windows.Forms.Panel panelParH;
        private System.Windows.Forms.Button btnAggiungiParH;
        private System.Windows.Forms.DataGridView dataGridParametriH;
        private System.Windows.Forms.ComboBox cmbParametriH;
        private System.Windows.Forms.Panel panelTopParD;
        private System.Windows.Forms.Panel panelTopParH;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.ContextMenuStrip contextMenuDataGrid;
        private System.Windows.Forms.ToolStripMenuItem modificareValoreToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cancellaParametroToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aggiungiNuovoToolStripMenuItem;
    }
}