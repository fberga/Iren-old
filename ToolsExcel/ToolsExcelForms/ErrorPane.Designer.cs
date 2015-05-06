namespace Iren.ToolsExcel.Forms
{
    partial class ErrorPane
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

        #region Codice generato da Progettazione componenti

        /// <summary> 
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare 
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Nodo0");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Nodo1");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("Nodo2");
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Nodo3");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Nodo4");
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("Nodo5");
            this.panelDescrizione = new System.Windows.Forms.Panel();
            this.lbTesto = new System.Windows.Forms.Label();
            this.lbTitolo = new System.Windows.Forms.Label();
            this.panelContent = new System.Windows.Forms.Panel();
            this.treeViewErrori = new System.Windows.Forms.TreeView();
            this.panelSeparator = new System.Windows.Forms.Panel();
            this.panelPadding = new System.Windows.Forms.Panel();
            this.panelDescrizione.SuspendLayout();
            this.panelContent.SuspendLayout();
            this.panelPadding.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelDescrizione
            // 
            this.panelDescrizione.Controls.Add(this.lbTesto);
            this.panelDescrizione.Controls.Add(this.lbTitolo);
            this.panelDescrizione.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelDescrizione.Location = new System.Drawing.Point(0, 0);
            this.panelDescrizione.Name = "panelDescrizione";
            this.panelDescrizione.Padding = new System.Windows.Forms.Padding(4, 1, 2, 0);
            this.panelDescrizione.Size = new System.Drawing.Size(382, 86);
            this.panelDescrizione.TabIndex = 0;
            // 
            // lbTesto
            // 
            this.lbTesto.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbTesto.Location = new System.Drawing.Point(4, 17);
            this.lbTesto.Name = "lbTesto";
            this.lbTesto.Size = new System.Drawing.Size(376, 69);
            this.lbTesto.TabIndex = 3;
            this.lbTesto.Text = "Il pannello contiene la lista di errori suddivisi per UP e per giorno.";
            // 
            // lbTitolo
            // 
            this.lbTitolo.AutoSize = true;
            this.lbTitolo.Dock = System.Windows.Forms.DockStyle.Top;
            this.lbTitolo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbTitolo.Location = new System.Drawing.Point(4, 1);
            this.lbTitolo.Margin = new System.Windows.Forms.Padding(5);
            this.lbTitolo.Name = "lbTitolo";
            this.lbTitolo.Size = new System.Drawing.Size(149, 16);
            this.lbTitolo.TabIndex = 2;
            this.lbTitolo.Text = "Pannello degli errori";
            // 
            // panelContent
            // 
            this.panelContent.Controls.Add(this.treeViewErrori);
            this.panelContent.Controls.Add(this.panelPadding);
            this.panelContent.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelContent.Location = new System.Drawing.Point(0, 86);
            this.panelContent.Name = "panelContent";
            this.panelContent.Size = new System.Drawing.Size(382, 341);
            this.panelContent.TabIndex = 1;
            // 
            // treeViewErrori
            // 
            this.treeViewErrori.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.treeViewErrori.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewErrori.Location = new System.Drawing.Point(0, 10);
            this.treeViewErrori.Name = "treeViewErrori";
            treeNode1.Name = "Nodo0";
            treeNode1.Text = "Nodo0";
            treeNode2.Name = "Nodo1";
            treeNode2.Text = "Nodo1";
            treeNode3.Name = "Nodo2";
            treeNode3.Text = "Nodo2";
            treeNode4.Name = "Nodo3";
            treeNode4.Text = "Nodo3";
            treeNode5.Name = "Nodo4";
            treeNode5.Text = "Nodo4";
            treeNode6.Name = "Nodo5";
            treeNode6.Text = "Nodo5";
            this.treeViewErrori.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3,
            treeNode4,
            treeNode5,
            treeNode6});
            this.treeViewErrori.Size = new System.Drawing.Size(382, 331);
            this.treeViewErrori.TabIndex = 0;
            // 
            // panelSeparator
            // 
            this.panelSeparator.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelSeparator.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelSeparator.Location = new System.Drawing.Point(4, 0);
            this.panelSeparator.Margin = new System.Windows.Forms.Padding(0);
            this.panelSeparator.Name = "panelSeparator";
            this.panelSeparator.Size = new System.Drawing.Size(374, 1);
            this.panelSeparator.TabIndex = 1;
            // 
            // panelPadding
            // 
            this.panelPadding.Controls.Add(this.panelSeparator);
            this.panelPadding.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelPadding.Location = new System.Drawing.Point(0, 0);
            this.panelPadding.Name = "panelPadding";
            this.panelPadding.Padding = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.panelPadding.Size = new System.Drawing.Size(382, 10);
            this.panelPadding.TabIndex = 2;
            // 
            // ErrorPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panelContent);
            this.Controls.Add(this.panelDescrizione);
            this.Name = "ErrorPane";
            this.Size = new System.Drawing.Size(382, 575);
            this.SizeChanged += new System.EventHandler(this.ErrorPane_SizeChanged);
            this.panelDescrizione.ResumeLayout(false);
            this.panelDescrizione.PerformLayout();
            this.panelContent.ResumeLayout(false);
            this.panelPadding.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelDescrizione;
        private System.Windows.Forms.Label lbTesto;
        private System.Windows.Forms.Label lbTitolo;
        private System.Windows.Forms.Panel panelContent;
        private System.Windows.Forms.TreeView treeViewErrori;
        private System.Windows.Forms.Panel panelSeparator;
        private System.Windows.Forms.Panel panelPadding;


    }
}
