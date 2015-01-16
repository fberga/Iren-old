namespace Iren.FrontOffice.Forms
{
    partial class frmRAMPE
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Liberare le risorse in uso.
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgModificaRampa = new System.Windows.Forms.DataGridView();
            this.cmbEntita = new System.Windows.Forms.ComboBox();
            this.dgVisualizzaRampe = new System.Windows.Forms.DataGridView();
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgModificaRampa)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgVisualizzaRampe)).BeginInit();
            this.SuspendLayout();
            // 
            // dgModificaRampa
            // 
            this.dgModificaRampa.AllowUserToAddRows = false;
            this.dgModificaRampa.AllowUserToDeleteRows = false;
            this.dgModificaRampa.AllowUserToResizeColumns = false;
            this.dgModificaRampa.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgModificaRampa.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgModificaRampa.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgModificaRampa.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgModificaRampa.Location = new System.Drawing.Point(14, 55);
            this.dgModificaRampa.Name = "dgModificaRampa";
            dataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.Black;
            this.dgModificaRampa.RowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dgModificaRampa.Size = new System.Drawing.Size(1003, 179);
            this.dgModificaRampa.TabIndex = 0;
            this.dgModificaRampa.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dgModificaRampa_CellBeginEdit);
            this.dgModificaRampa.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dgModificaRampa_CellPainting);
            this.dgModificaRampa.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgModificaRampa_CellValueChanged);
            this.dgModificaRampa.ColumnAdded += new System.Windows.Forms.DataGridViewColumnEventHandler(this.dgModificaRampa_ColumnAdded);
            this.dgModificaRampa.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgModificaRampa_CurrentCellDirtyStateChanged);
            // 
            // cmbEntita
            // 
            this.cmbEntita.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbEntita.FormattingEnabled = true;
            this.cmbEntita.Location = new System.Drawing.Point(14, 12);
            this.cmbEntita.Name = "cmbEntita";
            this.cmbEntita.Size = new System.Drawing.Size(384, 26);
            this.cmbEntita.TabIndex = 1;
            this.cmbEntita.SelectedIndexChanged += new System.EventHandler(this.cmbEntita_SelectedIndexChanged);
            // 
            // dgVisualizzaRampe
            // 
            this.dgVisualizzaRampe.AllowUserToAddRows = false;
            this.dgVisualizzaRampe.AllowUserToDeleteRows = false;
            this.dgVisualizzaRampe.AllowUserToResizeColumns = false;
            this.dgVisualizzaRampe.AllowUserToResizeRows = false;
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgVisualizzaRampe.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dgVisualizzaRampe.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgVisualizzaRampe.DefaultCellStyle = dataGridViewCellStyle5;
            this.dgVisualizzaRampe.Location = new System.Drawing.Point(14, 250);
            this.dgVisualizzaRampe.Name = "dgVisualizzaRampe";
            dataGridViewCellStyle6.SelectionForeColor = System.Drawing.Color.Black;
            this.dgVisualizzaRampe.RowsDefaultCellStyle = dataGridViewCellStyle6;
            this.dgVisualizzaRampe.Size = new System.Drawing.Size(1003, 179);
            this.dgVisualizzaRampe.TabIndex = 2;
            this.dgVisualizzaRampe.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.dgVisualizzaRampe_CellPainting);
            this.dgVisualizzaRampe.ColumnAdded += new System.Windows.Forms.DataGridViewColumnEventHandler(this.dgVisualizzaRampe_ColumnAdded);
            // 
            // btnApplica
            // 
            this.btnApplica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplica.Location = new System.Drawing.Point(783, 445);
            this.btnApplica.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnApplica.Name = "btnApplica";
            this.btnApplica.Size = new System.Drawing.Size(113, 49);
            this.btnApplica.TabIndex = 4;
            this.btnApplica.Text = "Applica";
            this.btnApplica.UseVisualStyleBackColor = true;
            this.btnApplica.Click += new System.EventHandler(this.btnApplica_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(904, 445);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 49);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);
            // 
            // frmRAMPE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnAnnulla;
            this.ClientSize = new System.Drawing.Size(1028, 508);
            this.Controls.Add(this.btnAnnulla);
            this.Controls.Add(this.btnApplica);
            this.Controls.Add(this.dgVisualizzaRampe);
            this.Controls.Add(this.cmbEntita);
            this.Controls.Add(this.dgModificaRampa);
            this.Name = "frmRAMPE";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "Rampe";
            this.Load += new System.EventHandler(this.frmRAMPE_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgModificaRampa)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgVisualizzaRampe)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgModificaRampa;
        private System.Windows.Forms.ComboBox cmbEntita;
        private System.Windows.Forms.DataGridView dgVisualizzaRampe;
        private System.Windows.Forms.Button btnApplica;
        private System.Windows.Forms.Button btnAnnulla;


    }
}