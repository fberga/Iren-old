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
            this.cmbEntita = new System.Windows.Forms.ComboBox();
            this.tabParametri = new System.Windows.Forms.TabControl();
            this.tabPageParD = new System.Windows.Forms.TabPage();
            this.tabPageParH = new System.Windows.Forms.TabPage();
            this.panelParD = new System.Windows.Forms.Panel();
            this.btnRimuoviParD = new System.Windows.Forms.Button();
            this.btnAggiungiParD = new System.Windows.Forms.Button();
            this.dataGridParametriD = new System.Windows.Forms.DataGridView();
            this.cmbParametriD = new System.Windows.Forms.ComboBox();
            this.panelParH = new System.Windows.Forms.Panel();
            this.btnRimuoviParH = new System.Windows.Forms.Button();
            this.btnAggiungiParH = new System.Windows.Forms.Button();
            this.dataGridParametriH = new System.Windows.Forms.DataGridView();
            this.cmbParametriH = new System.Windows.Forms.ComboBox();
            this.tabParametri.SuspendLayout();
            this.tabPageParD.SuspendLayout();
            this.tabPageParH.SuspendLayout();
            this.panelParD.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriD)).BeginInit();
            this.panelParH.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriH)).BeginInit();
            this.SuspendLayout();
            // 
            // cmbEntita
            // 
            this.cmbEntita.FormattingEnabled = true;
            this.cmbEntita.Location = new System.Drawing.Point(12, 12);
            this.cmbEntita.Name = "cmbEntita";
            this.cmbEntita.Size = new System.Drawing.Size(294, 28);
            this.cmbEntita.TabIndex = 0;
            this.cmbEntita.SelectedIndexChanged += new System.EventHandler(this.cmbEntita_SelectedIndexChanged);
            // 
            // tabParametri
            // 
            this.tabParametri.Controls.Add(this.tabPageParD);
            this.tabParametri.Controls.Add(this.tabPageParH);
            this.tabParametri.Location = new System.Drawing.Point(12, 46);
            this.tabParametri.Name = "tabParametri";
            this.tabParametri.SelectedIndex = 0;
            this.tabParametri.Size = new System.Drawing.Size(1208, 434);
            this.tabParametri.TabIndex = 1;
            // 
            // tabPageParD
            // 
            this.tabPageParD.Controls.Add(this.panelParD);
            this.tabPageParD.Controls.Add(this.dataGridParametriD);
            this.tabPageParD.Controls.Add(this.cmbParametriD);
            this.tabPageParD.Location = new System.Drawing.Point(4, 29);
            this.tabPageParD.Name = "tabPageParD";
            this.tabPageParD.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageParD.Size = new System.Drawing.Size(1200, 401);
            this.tabPageParD.TabIndex = 0;
            this.tabPageParD.Text = "Parametri Giornalieri";
            this.tabPageParD.UseVisualStyleBackColor = true;
            // 
            // tabPageParH
            // 
            this.tabPageParH.Controls.Add(this.panelParH);
            this.tabPageParH.Controls.Add(this.dataGridParametriH);
            this.tabPageParH.Controls.Add(this.cmbParametriH);
            this.tabPageParH.Location = new System.Drawing.Point(4, 29);
            this.tabPageParH.Name = "tabPageParH";
            this.tabPageParH.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageParH.Size = new System.Drawing.Size(1200, 401);
            this.tabPageParH.TabIndex = 1;
            this.tabPageParH.Text = "Parametri Orari";
            this.tabPageParH.UseVisualStyleBackColor = true;
            // 
            // panelParD
            // 
            this.panelParD.Controls.Add(this.btnRimuoviParD);
            this.panelParD.Controls.Add(this.btnAggiungiParD);
            this.panelParD.Location = new System.Drawing.Point(6, 296);
            this.panelParD.Name = "panelParD";
            this.panelParD.Size = new System.Drawing.Size(1188, 46);
            this.panelParD.TabIndex = 7;
            // 
            // btnRimuoviParD
            // 
            this.btnRimuoviParD.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnRimuoviParD.Location = new System.Drawing.Point(1096, 0);
            this.btnRimuoviParD.Name = "btnRimuoviParD";
            this.btnRimuoviParD.Size = new System.Drawing.Size(46, 46);
            this.btnRimuoviParD.TabIndex = 1;
            this.btnRimuoviParD.Text = "-";
            this.btnRimuoviParD.UseVisualStyleBackColor = true;
            // 
            // btnAggiungiParD
            // 
            this.btnAggiungiParD.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAggiungiParD.Location = new System.Drawing.Point(1142, 0);
            this.btnAggiungiParD.Name = "btnAggiungiParD";
            this.btnAggiungiParD.Size = new System.Drawing.Size(46, 46);
            this.btnAggiungiParD.TabIndex = 0;
            this.btnAggiungiParD.Text = "+";
            this.btnAggiungiParD.UseVisualStyleBackColor = true;
            // 
            // dataGridParametriD
            // 
            this.dataGridParametriD.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridParametriD.Location = new System.Drawing.Point(6, 40);
            this.dataGridParametriD.Name = "dataGridParametriD";
            this.dataGridParametriD.Size = new System.Drawing.Size(1188, 256);
            this.dataGridParametriD.TabIndex = 6;
            // 
            // cmbParametriD
            // 
            this.cmbParametriD.FormattingEnabled = true;
            this.cmbParametriD.Location = new System.Drawing.Point(6, 6);
            this.cmbParametriD.Name = "cmbParametriD";
            this.cmbParametriD.Size = new System.Drawing.Size(294, 28);
            this.cmbParametriD.TabIndex = 5;
            this.cmbParametriD.SelectedIndexChanged += new System.EventHandler(this.cmbParametriD_SelectedIndexChanged);
            // 
            // panelParH
            // 
            this.panelParH.Controls.Add(this.btnRimuoviParH);
            this.panelParH.Controls.Add(this.btnAggiungiParH);
            this.panelParH.Location = new System.Drawing.Point(6, 296);
            this.panelParH.Name = "panelParH";
            this.panelParH.Size = new System.Drawing.Size(1188, 46);
            this.panelParH.TabIndex = 7;
            // 
            // btnRimuoviParH
            // 
            this.btnRimuoviParH.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnRimuoviParH.Location = new System.Drawing.Point(1096, 0);
            this.btnRimuoviParH.Name = "btnRimuoviParH";
            this.btnRimuoviParH.Size = new System.Drawing.Size(46, 46);
            this.btnRimuoviParH.TabIndex = 1;
            this.btnRimuoviParH.Text = "-";
            this.btnRimuoviParH.UseVisualStyleBackColor = true;
            // 
            // btnAggiungiParH
            // 
            this.btnAggiungiParH.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAggiungiParH.Location = new System.Drawing.Point(1142, 0);
            this.btnAggiungiParH.Name = "btnAggiungiParH";
            this.btnAggiungiParH.Size = new System.Drawing.Size(46, 46);
            this.btnAggiungiParH.TabIndex = 0;
            this.btnAggiungiParH.Text = "+";
            this.btnAggiungiParH.UseVisualStyleBackColor = true;
            // 
            // dataGridParametriH
            // 
            this.dataGridParametriH.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridParametriH.Location = new System.Drawing.Point(6, 40);
            this.dataGridParametriH.Name = "dataGridParametriH";
            this.dataGridParametriH.Size = new System.Drawing.Size(1188, 256);
            this.dataGridParametriH.TabIndex = 6;
            // 
            // cmbParametriH
            // 
            this.cmbParametriH.FormattingEnabled = true;
            this.cmbParametriH.Location = new System.Drawing.Point(6, 6);
            this.cmbParametriH.Name = "cmbParametriH";
            this.cmbParametriH.Size = new System.Drawing.Size(294, 28);
            this.cmbParametriH.TabIndex = 5;
            // 
            // FormModificaParametri
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1232, 492);
            this.Controls.Add(this.tabParametri);
            this.Controls.Add(this.cmbEntita);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormModificaParametri";
            this.Text = "FormModificaParametri";
            this.tabParametri.ResumeLayout(false);
            this.tabPageParD.ResumeLayout(false);
            this.tabPageParH.ResumeLayout(false);
            this.panelParD.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriD)).EndInit();
            this.panelParH.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametriH)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbEntita;
        private System.Windows.Forms.TabControl tabParametri;
        private System.Windows.Forms.TabPage tabPageParD;
        private System.Windows.Forms.Panel panelParD;
        private System.Windows.Forms.Button btnRimuoviParD;
        private System.Windows.Forms.Button btnAggiungiParD;
        private System.Windows.Forms.DataGridView dataGridParametriD;
        private System.Windows.Forms.ComboBox cmbParametriD;
        private System.Windows.Forms.TabPage tabPageParH;
        private System.Windows.Forms.Panel panelParH;
        private System.Windows.Forms.Button btnRimuoviParH;
        private System.Windows.Forms.Button btnAggiungiParH;
        private System.Windows.Forms.DataGridView dataGridParametriH;
        private System.Windows.Forms.ComboBox cmbParametriH;
    }
}