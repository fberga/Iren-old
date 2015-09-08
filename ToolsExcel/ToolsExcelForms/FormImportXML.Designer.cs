namespace Iren.ToolsExcel.Forms
{
    partial class FormImportXML
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.dataGridFileXML = new System.Windows.Forms.DataGridView();
            this.openFileXMLImport = new System.Windows.Forms.OpenFileDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnImporta = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.btnApri = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridFileXML)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.dataGridFileXML, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 400F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 54F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.Size = new System.Drawing.Size(1072, 475);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // dataGridFileXML
            // 
            this.dataGridFileXML.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridFileXML.Location = new System.Drawing.Point(3, 23);
            this.dataGridFileXML.Name = "dataGridFileXML";
            this.dataGridFileXML.Size = new System.Drawing.Size(982, 362);
            this.dataGridFileXML.TabIndex = 0;
            // 
            // openFileXMLImport
            // 
            this.openFileXMLImport.FileName = "openFileDialog1";
            this.openFileXMLImport.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileXMLImport_FileOk);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnApri);
            this.panel1.Controls.Add(this.btnImporta);
            this.panel1.Controls.Add(this.btnAnnulla);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(3, 423);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1066, 48);
            this.panel1.TabIndex = 2;
            // 
            // btnImporta
            // 
            this.btnImporta.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnImporta.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImporta.Location = new System.Drawing.Point(840, 0);
            this.btnImporta.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnImporta.Name = "btnImporta";
            this.btnImporta.Size = new System.Drawing.Size(113, 48);
            this.btnImporta.TabIndex = 6;
            this.btnImporta.Text = "Importa dati";
            this.btnImporta.UseVisualStyleBackColor = true;
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(953, 0);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 48);
            this.btnAnnulla.TabIndex = 7;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // btnApri
            // 
            this.btnApri.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApri.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApri.Location = new System.Drawing.Point(727, 0);
            this.btnApri.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnApri.Name = "btnApri";
            this.btnApri.Size = new System.Drawing.Size(113, 48);
            this.btnApri.TabIndex = 8;
            this.btnApri.Text = "Apri file";
            this.btnApri.UseVisualStyleBackColor = true;
            this.btnApri.Click += new System.EventHandler(this.btnApri_Click);
            // 
            // FormImportXML
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1072, 475);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "FormImportXML";
            this.Text = "FormImportXML";
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridFileXML)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView dataGridFileXML;
        private System.Windows.Forms.OpenFileDialog openFileXMLImport;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnImporta;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Button btnApri;
    }
}