namespace Iren.ToolsExcel
{
    partial class FormSettings
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
            this.dtDataAttiva = new System.Windows.Forms.DateTimePicker();
            this.lbDataAttiva = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // dtDataAttiva
            // 
            this.dtDataAttiva.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtDataAttiva.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtDataAttiva.Location = new System.Drawing.Point(106, 3);
            this.dtDataAttiva.Name = "dtDataAttiva";
            this.dtDataAttiva.Size = new System.Drawing.Size(127, 26);
            this.dtDataAttiva.TabIndex = 0;
            // 
            // lbDataAttiva
            // 
            this.lbDataAttiva.AutoSize = true;
            this.lbDataAttiva.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbDataAttiva.Location = new System.Drawing.Point(12, 8);
            this.lbDataAttiva.Name = "lbDataAttiva";
            this.lbDataAttiva.Size = new System.Drawing.Size(88, 20);
            this.lbDataAttiva.TabIndex = 1;
            this.lbDataAttiva.Text = "Data Attiva";
            // 
            // SettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(825, 328);
            this.Controls.Add(this.lbDataAttiva);
            this.Controls.Add(this.dtDataAttiva);
            this.Name = "SettingsForm";
            this.Text = "SettingsForm";
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dtDataAttiva;
        private System.Windows.Forms.Label lbDataAttiva;
    }
}