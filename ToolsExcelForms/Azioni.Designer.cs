namespace Iren.FrontOffice.Forms
{
    partial class frmAZIONI
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
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.panelTop = new System.Windows.Forms.Panel();
            this.panelMercati = new System.Windows.Forms.Panel();
            this.panelGiorni = new System.Windows.Forms.Panel();
            this.panelCentrale = new System.Windows.Forms.Panel();
            this.panelUP = new System.Windows.Forms.Panel();
            this.treeViewUP = new System.Windows.Forms.TreeView();
            this.panelCategorie = new System.Windows.Forms.Panel();
            this.treeViewCategorie = new System.Windows.Forms.TreeView();
            this.panelAzioni = new System.Windows.Forms.Panel();
            this.treeViewAzioni = new System.Windows.Forms.TreeView();
            this.panelButtons.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.panelCentrale.SuspendLayout();
            this.panelUP.SuspendLayout();
            this.panelCategorie.SuspendLayout();
            this.panelAzioni.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnApplica);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(5, 440);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 5, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(1050, 53);
            this.panelButtons.TabIndex = 12;
            // 
            // btnApplica
            // 
            this.btnApplica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApplica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplica.Location = new System.Drawing.Point(824, 5);
            this.btnApplica.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnApplica.Name = "btnApplica";
            this.btnApplica.Size = new System.Drawing.Size(113, 48);
            this.btnApplica.TabIndex = 4;
            this.btnApplica.Text = "Applica";
            this.btnApplica.UseVisualStyleBackColor = true;
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(937, 5);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 48);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.panelMercati);
            this.panelTop.Controls.Add(this.panelGiorni);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(5, 5);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(1050, 53);
            this.panelTop.TabIndex = 13;
            // 
            // panelMercati
            // 
            this.panelMercati.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelMercati.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panelMercati.Location = new System.Drawing.Point(350, 0);
            this.panelMercati.Name = "panelMercati";
            this.panelMercati.Size = new System.Drawing.Size(700, 53);
            this.panelMercati.TabIndex = 3;
            // 
            // panelGiorni
            // 
            this.panelGiorni.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelGiorni.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panelGiorni.Location = new System.Drawing.Point(0, 0);
            this.panelGiorni.Name = "panelGiorni";
            this.panelGiorni.Size = new System.Drawing.Size(350, 53);
            this.panelGiorni.TabIndex = 2;
            // 
            // panelCentrale
            // 
            this.panelCentrale.Controls.Add(this.panelUP);
            this.panelCentrale.Controls.Add(this.panelCategorie);
            this.panelCentrale.Controls.Add(this.panelAzioni);
            this.panelCentrale.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelCentrale.Location = new System.Drawing.Point(5, 58);
            this.panelCentrale.Name = "panelCentrale";
            this.panelCentrale.Padding = new System.Windows.Forms.Padding(0, 5, 0, 5);
            this.panelCentrale.Size = new System.Drawing.Size(1050, 382);
            this.panelCentrale.TabIndex = 14;
            // 
            // panelUP
            // 
            this.panelUP.Controls.Add(this.treeViewUP);
            this.panelUP.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelUP.Location = new System.Drawing.Point(700, 5);
            this.panelUP.Name = "panelUP";
            this.panelUP.Size = new System.Drawing.Size(350, 372);
            this.panelUP.TabIndex = 6;
            // 
            // treeViewUP
            // 
            this.treeViewUP.CheckBoxes = true;
            this.treeViewUP.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewUP.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewUP.Location = new System.Drawing.Point(0, 0);
            this.treeViewUP.Name = "treeViewUP";
            this.treeViewUP.ShowNodeToolTips = true;
            this.treeViewUP.ShowPlusMinus = false;
            this.treeViewUP.ShowRootLines = false;
            this.treeViewUP.Size = new System.Drawing.Size(350, 372);
            this.treeViewUP.TabIndex = 1;
            // 
            // panelCategorie
            // 
            this.panelCategorie.Controls.Add(this.treeViewCategorie);
            this.panelCategorie.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelCategorie.Location = new System.Drawing.Point(350, 5);
            this.panelCategorie.Name = "panelCategorie";
            this.panelCategorie.Size = new System.Drawing.Size(350, 372);
            this.panelCategorie.TabIndex = 5;
            // 
            // treeViewCategorie
            // 
            this.treeViewCategorie.CheckBoxes = true;
            this.treeViewCategorie.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewCategorie.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewCategorie.Location = new System.Drawing.Point(0, 0);
            this.treeViewCategorie.Name = "treeViewCategorie";
            this.treeViewCategorie.ShowNodeToolTips = true;
            this.treeViewCategorie.ShowPlusMinus = false;
            this.treeViewCategorie.ShowRootLines = false;
            this.treeViewCategorie.Size = new System.Drawing.Size(350, 372);
            this.treeViewCategorie.TabIndex = 1;
            this.treeViewCategorie.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeView_AfterCheck);
            this.treeViewCategorie.BeforeCollapse += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeView_BeforeCollapse);
            // 
            // panelAzioni
            // 
            this.panelAzioni.Controls.Add(this.treeViewAzioni);
            this.panelAzioni.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelAzioni.Location = new System.Drawing.Point(0, 5);
            this.panelAzioni.Name = "panelAzioni";
            this.panelAzioni.Size = new System.Drawing.Size(350, 372);
            this.panelAzioni.TabIndex = 4;
            // 
            // treeViewAzioni
            // 
            this.treeViewAzioni.CheckBoxes = true;
            this.treeViewAzioni.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewAzioni.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewAzioni.Location = new System.Drawing.Point(0, 0);
            this.treeViewAzioni.Name = "treeViewAzioni";
            this.treeViewAzioni.ShowPlusMinus = false;
            this.treeViewAzioni.ShowRootLines = false;
            this.treeViewAzioni.Size = new System.Drawing.Size(350, 372);
            this.treeViewAzioni.TabIndex = 0;
            this.treeViewAzioni.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeView_AfterCheck);
            this.treeViewAzioni.BeforeCollapse += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeView_BeforeCollapse);
            // 
            // frmAZIONI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1060, 498);
            this.Controls.Add(this.panelCentrale);
            this.Controls.Add(this.panelTop);
            this.Controls.Add(this.panelButtons);
            this.Name = "frmAZIONI";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.Text = "Azioni";
            this.Load += new System.EventHandler(this.frmAZIONI_Load);
            this.panelButtons.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.panelCentrale.ResumeLayout(false);
            this.panelUP.ResumeLayout(false);
            this.panelCategorie.ResumeLayout(false);
            this.panelAzioni.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnApplica;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.Panel panelMercati;
        private System.Windows.Forms.Panel panelGiorni;
        private System.Windows.Forms.Panel panelCentrale;
        private System.Windows.Forms.Panel panelUP;
        private System.Windows.Forms.Panel panelCategorie;
        private System.Windows.Forms.Panel panelAzioni;
        private System.Windows.Forms.TreeView treeViewUP;
        private System.Windows.Forms.TreeView treeViewCategorie;
        private System.Windows.Forms.TreeView treeViewAzioni;
    }
}