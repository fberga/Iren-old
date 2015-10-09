namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    partial class AssegnaFunzioni
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
            this.treeViewNotUtilized = new System.Windows.Forms.TreeView();
            this.treeViewUtilized = new System.Windows.Forms.TreeView();
            this.btnRimuovi = new System.Windows.Forms.Button();
            this.btnAggiungi = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.panelContent = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnSalva = new System.Windows.Forms.Button();
            this.panelContent.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeViewNotUtilized
            // 
            this.treeViewNotUtilized.Location = new System.Drawing.Point(4, 29);
            this.treeViewNotUtilized.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.treeViewNotUtilized.Name = "treeViewNotUtilized";
            this.treeViewNotUtilized.Size = new System.Drawing.Size(272, 493);
            this.treeViewNotUtilized.TabIndex = 2;
            // 
            // treeViewUtilized
            // 
            this.treeViewUtilized.Location = new System.Drawing.Point(324, 29);
            this.treeViewUtilized.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.treeViewUtilized.Name = "treeViewUtilized";
            this.treeViewUtilized.Size = new System.Drawing.Size(272, 493);
            this.treeViewUtilized.TabIndex = 3;
            // 
            // btnRimuovi
            // 
            this.btnRimuovi.BackgroundImage = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.leftArrow;
            this.btnRimuovi.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnRimuovi.FlatAppearance.BorderSize = 0;
            this.btnRimuovi.Location = new System.Drawing.Point(284, 279);
            this.btnRimuovi.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnRimuovi.Name = "btnRimuovi";
            this.btnRimuovi.Size = new System.Drawing.Size(32, 30);
            this.btnRimuovi.TabIndex = 4;
            this.btnRimuovi.UseVisualStyleBackColor = true;
            this.btnRimuovi.Click += new System.EventHandler(this.RimuoviFunzione_Click);
            // 
            // btnAggiungi
            // 
            this.btnAggiungi.BackgroundImage = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.rightArrow;
            this.btnAggiungi.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAggiungi.FlatAppearance.BorderSize = 0;
            this.btnAggiungi.Location = new System.Drawing.Point(284, 241);
            this.btnAggiungi.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAggiungi.Name = "btnAggiungi";
            this.btnAggiungi.Size = new System.Drawing.Size(32, 30);
            this.btnAggiungi.TabIndex = 5;
            this.btnAggiungi.UseVisualStyleBackColor = true;
            this.btnAggiungi.Click += new System.EventHandler(this.AggiungiFunzione_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(1, 9);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "Funzioni disponibili";
            // 
            // panelContent
            // 
            this.panelContent.Controls.Add(this.panelBottom);
            this.panelContent.Controls.Add(this.label2);
            this.panelContent.Controls.Add(this.label1);
            this.panelContent.Controls.Add(this.treeViewNotUtilized);
            this.panelContent.Controls.Add(this.btnAggiungi);
            this.panelContent.Controls.Add(this.treeViewUtilized);
            this.panelContent.Controls.Add(this.btnRimuovi);
            this.panelContent.Location = new System.Drawing.Point(0, 0);
            this.panelContent.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.panelContent.Name = "panelContent";
            this.panelContent.Size = new System.Drawing.Size(596, 572);
            this.panelContent.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(321, 9);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 16);
            this.label2.TabIndex = 7;
            this.label2.Text = "Funzioni assegnate";
            // 
            // panelBottom
            // 
            this.panelBottom.Controls.Add(this.btnSalva);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelBottom.Location = new System.Drawing.Point(0, 524);
            this.panelBottom.Margin = new System.Windows.Forms.Padding(0);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(596, 48);
            this.panelBottom.TabIndex = 15;
            // 
            // btnSalva
            // 
            this.btnSalva.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnSalva.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSalva.Location = new System.Drawing.Point(483, 0);
            this.btnSalva.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSalva.Name = "btnSalva";
            this.btnSalva.Size = new System.Drawing.Size(113, 48);
            this.btnSalva.TabIndex = 8;
            this.btnSalva.Text = "Salva";
            this.btnSalva.UseVisualStyleBackColor = true;
            // 
            // AssegnaFunzioni
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(805, 609);
            this.Controls.Add(this.panelContent);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "AssegnaFunzioni";
            this.Text = "AssegnaFunzioni";
            this.panelContent.ResumeLayout(false);
            this.panelContent.PerformLayout();
            this.panelBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeViewNotUtilized;
        private System.Windows.Forms.TreeView treeViewUtilized;
        private System.Windows.Forms.Button btnRimuovi;
        private System.Windows.Forms.Button btnAggiungi;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panelContent;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnSalva;
    }
}