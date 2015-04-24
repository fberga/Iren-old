namespace Iren.ToolsExcel.Forms
{
    partial class LoaderScreen
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
            this.lbCaricamento = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lbText = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // lbCaricamento
            // 
            this.lbCaricamento.AutoSize = true;
            this.lbCaricamento.BackColor = System.Drawing.Color.Black;
            this.lbCaricamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbCaricamento.ForeColor = System.Drawing.SystemColors.Control;
            this.lbCaricamento.Location = new System.Drawing.Point(182, 72);
            this.lbCaricamento.Name = "lbCaricamento";
            this.lbCaricamento.Size = new System.Drawing.Size(124, 18);
            this.lbCaricamento.TabIndex = 0;
            this.lbCaricamento.Text = "Caricamento ...";
            this.lbCaricamento.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Image = global::Iren.ToolsExcel.Forms.Properties.Resources.loader;
            this.pictureBox1.Location = new System.Drawing.Point(1, 1);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(481, 180);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // lbText
            // 
            this.lbText.AutoSize = true;
            this.lbText.BackColor = System.Drawing.Color.Black;
            this.lbText.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbText.ForeColor = System.Drawing.SystemColors.Control;
            this.lbText.Location = new System.Drawing.Point(218, 100);
            this.lbText.Name = "lbText";
            this.lbText.Size = new System.Drawing.Size(46, 18);
            this.lbText.TabIndex = 2;
            this.lbText.Text = "label1";
            this.lbText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lbText.SizeChanged += new System.EventHandler(this.lbText_SizeChanged);
            // 
            // LoaderScreen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.ClientSize = new System.Drawing.Size(483, 182);
            this.Controls.Add(this.lbText);
            this.Controls.Add(this.lbCaricamento);
            this.Controls.Add(this.pictureBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "LoaderScreen";
            this.Padding = new System.Windows.Forms.Padding(1);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "LoaderScreen";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.LoaderScreen_Load);
            this.VisibleChanged += new System.EventHandler(this.LoaderScreen_VisibleChanged);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbCaricamento;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lbText;

    }
}