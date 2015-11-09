namespace Iren.PSO.Launcher
{
    partial class Launcher
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Launcher));
            this.IconTray = new System.Windows.Forms.NotifyIcon(this.components);
            this.menuIconTray = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.menuLayout = new System.Windows.Forms.TableLayoutPanel();
            this.SuspendLayout();
            // 
            // IconTray
            // 
            this.IconTray.ContextMenuStrip = this.menuIconTray;
            this.IconTray.Icon = ((System.Drawing.Icon)(resources.GetObject("IconTray.Icon")));
            this.IconTray.Text = "PSO";
            this.IconTray.Visible = true;
            this.IconTray.MouseClick += new System.Windows.Forms.MouseEventHandler(this.IconTray_MouseClick);
            this.IconTray.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.IconTray_MouseDoubleClick);
            // 
            // menuIconTray
            // 
            this.menuIconTray.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuIconTray.ImageScalingSize = new System.Drawing.Size(28, 28);
            this.menuIconTray.Name = "Menu";
            this.menuIconTray.Size = new System.Drawing.Size(61, 4);
            this.menuIconTray.Text = "PSO";
            // 
            // menuLayout
            // 
            this.menuLayout.AutoSize = true;
            this.menuLayout.ColumnCount = 1;
            this.menuLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.menuLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.menuLayout.Location = new System.Drawing.Point(0, 0);
            this.menuLayout.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.menuLayout.Name = "menuLayout";
            this.menuLayout.RowCount = 1;
            this.menuLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.menuLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.menuLayout.Size = new System.Drawing.Size(85, 408);
            this.menuLayout.TabIndex = 1;
            // 
            // Launcher
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(117, 441);
            this.Controls.Add(this.menuLayout);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Launcher";
            this.Opacity = 0D;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PSO";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Launcher_FormClosing);
            this.Load += new System.EventHandler(this.Launcher_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.NotifyIcon IconTray;
        private System.Windows.Forms.ContextMenuStrip menuIconTray;
        private System.Windows.Forms.TableLayoutPanel menuLayout;

    }
}

