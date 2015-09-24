namespace PSOLauncher
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.IconTray = new System.Windows.Forms.NotifyIcon(this.components);
            this.menuTrayIcon = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.previsioneGas = new System.Windows.Forms.ToolStripMenuItem();
            this.unitComm = new System.Windows.Forms.ToolStripMenuItem();
            this.prezziMSD = new System.Windows.Forms.ToolStripMenuItem();
            this.validazioneTL = new System.Windows.Forms.ToolStripMenuItem();
            this.previsioneCT = new System.Windows.Forms.ToolStripMenuItem();
            this.progrImp = new System.Windows.Forms.ToolStripMenuItem();
            this.offerteMGP = new System.Windows.Forms.ToolStripMenuItem();
            this.offerteMSD = new System.Windows.Forms.ToolStripMenuItem();
            this.offerteMB = new System.Windows.Forms.ToolStripMenuItem();
            this.invioProgrammi = new System.Windows.Forms.ToolStripMenuItem();
            this.sisCom = new System.Windows.Forms.ToolStripMenuItem();
            this.btnPrevisioneGas = new System.Windows.Forms.Button();
            this.btnUnitComm = new System.Windows.Forms.Button();
            this.btnPrezziMSD = new System.Windows.Forms.Button();
            this.btnValidazioneTL = new System.Windows.Forms.Button();
            this.btnPrevisioneCT = new System.Windows.Forms.Button();
            this.btnProgrImp = new System.Windows.Forms.Button();
            this.btnOfferteMGP = new System.Windows.Forms.Button();
            this.btnOfferteMSD = new System.Windows.Forms.Button();
            this.btnOfferteMB = new System.Windows.Forms.Button();
            this.btnInvioProgrammi = new System.Windows.Forms.Button();
            this.btnSisCom = new System.Windows.Forms.Button();
            this.menuTrayIcon.SuspendLayout();
            this.SuspendLayout();
            // 
            // IconTray
            // 
            this.IconTray.ContextMenuStrip = this.menuTrayIcon;
            this.IconTray.Icon = ((System.Drawing.Icon)(resources.GetObject("IconTray.Icon")));
            this.IconTray.Text = "PSO";
            this.IconTray.Visible = true;
            // 
            // menuTrayIcon
            // 
            this.menuTrayIcon.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuTrayIcon.ImageScalingSize = new System.Drawing.Size(28, 28);
            this.menuTrayIcon.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.previsioneGas,
            this.unitComm,
            this.prezziMSD,
            this.validazioneTL,
            this.previsioneCT,
            this.progrImp,
            this.offerteMGP,
            this.offerteMSD,
            this.offerteMB,
            this.invioProgrammi,
            this.sisCom});
            this.menuTrayIcon.Name = "Menu";
            this.menuTrayIcon.Size = new System.Drawing.Size(303, 378);
            this.menuTrayIcon.Text = "PSO";
            // 
            // previsioneGas
            // 
            this.previsioneGas.Image = global::PSOLauncher.Properties.Resources.gas_icon;
            this.previsioneGas.Name = "previsioneGas";
            this.previsioneGas.Size = new System.Drawing.Size(302, 34);
            this.previsioneGas.Text = "Previsione GAS";
            // 
            // unitComm
            // 
            this.unitComm.Image = global::PSOLauncher.Properties.Resources.unitComm_icon;
            this.unitComm.Name = "unitComm";
            this.unitComm.Size = new System.Drawing.Size(302, 34);
            this.unitComm.Text = "Unit Commitment";
            // 
            // prezziMSD
            // 
            this.prezziMSD.Image = global::PSOLauncher.Properties.Resources.prezziMSD_icon;
            this.prezziMSD.Name = "prezziMSD";
            this.prezziMSD.Size = new System.Drawing.Size(302, 34);
            this.prezziMSD.Text = "Prezzi MSD";
            // 
            // validazioneTL
            // 
            this.validazioneTL.Image = global::PSOLauncher.Properties.Resources.validazioneTL_icon;
            this.validazioneTL.Name = "validazioneTL";
            this.validazioneTL.Size = new System.Drawing.Size(302, 34);
            this.validazioneTL.Text = "Validazione Teleriscaldamento";
            // 
            // previsioneCT
            // 
            this.previsioneCT.Image = global::PSOLauncher.Properties.Resources.previsioneCT_icon;
            this.previsioneCT.Name = "previsioneCT";
            this.previsioneCT.Size = new System.Drawing.Size(302, 34);
            this.previsioneCT.Text = "Previsione Carico Termico";
            // 
            // progrImp
            // 
            this.progrImp.Image = global::PSOLauncher.Properties.Resources.progrImpianti_icon;
            this.progrImp.Name = "progrImp";
            this.progrImp.Size = new System.Drawing.Size(302, 34);
            this.progrImp.Text = "Programmazione Impianti";
            // 
            // offerteMGP
            // 
            this.offerteMGP.Image = global::PSOLauncher.Properties.Resources.offerteMGP_icon;
            this.offerteMGP.Name = "offerteMGP";
            this.offerteMGP.Size = new System.Drawing.Size(302, 34);
            this.offerteMGP.Text = "Offerte MGP";
            // 
            // offerteMSD
            // 
            this.offerteMSD.Image = global::PSOLauncher.Properties.Resources.offerteMSD_icon;
            this.offerteMSD.Name = "offerteMSD";
            this.offerteMSD.Size = new System.Drawing.Size(302, 34);
            this.offerteMSD.Text = "Offerte MSD";
            // 
            // offerteMB
            // 
            this.offerteMB.Image = global::PSOLauncher.Properties.Resources.offerteMB_icon;
            this.offerteMB.Name = "offerteMB";
            this.offerteMB.Size = new System.Drawing.Size(302, 34);
            this.offerteMB.Text = "Offerte MB";
            // 
            // invioProgrammi
            // 
            this.invioProgrammi.Image = global::PSOLauncher.Properties.Resources.invioProgrammi_icon;
            this.invioProgrammi.Name = "invioProgrammi";
            this.invioProgrammi.Size = new System.Drawing.Size(302, 34);
            this.invioProgrammi.Text = "Invio Programmi";
            // 
            // sisCom
            // 
            this.sisCom.Image = global::PSOLauncher.Properties.Resources.sisCom_icon;
            this.sisCom.Name = "sisCom";
            this.sisCom.Size = new System.Drawing.Size(302, 34);
            this.sisCom.Text = "Sistema Comandi";
            // 
            // btnPrevisioneGas
            // 
            this.btnPrevisioneGas.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnPrevisioneGas.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnPrevisioneGas.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnPrevisioneGas.FlatAppearance.BorderSize = 0;
            this.btnPrevisioneGas.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPrevisioneGas.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPrevisioneGas.Location = new System.Drawing.Point(0, 0);
            this.btnPrevisioneGas.Margin = new System.Windows.Forms.Padding(10);
            this.btnPrevisioneGas.Name = "btnPrevisioneGas";
            this.btnPrevisioneGas.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnPrevisioneGas.Size = new System.Drawing.Size(274, 34);
            this.btnPrevisioneGas.TabIndex = 1;
            this.btnPrevisioneGas.Text = "button1";
            this.btnPrevisioneGas.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPrevisioneGas.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnPrevisioneGas.UseVisualStyleBackColor = true;
            // 
            // btnUnitComm
            // 
            this.btnUnitComm.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnUnitComm.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnUnitComm.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnUnitComm.FlatAppearance.BorderSize = 0;
            this.btnUnitComm.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnUnitComm.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUnitComm.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnUnitComm.Location = new System.Drawing.Point(0, 34);
            this.btnUnitComm.Margin = new System.Windows.Forms.Padding(10);
            this.btnUnitComm.Name = "btnUnitComm";
            this.btnUnitComm.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnUnitComm.Size = new System.Drawing.Size(274, 34);
            this.btnUnitComm.TabIndex = 2;
            this.btnUnitComm.Text = "button2";
            this.btnUnitComm.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnUnitComm.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnUnitComm.UseVisualStyleBackColor = true;
            // 
            // btnPrezziMSD
            // 
            this.btnPrezziMSD.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnPrezziMSD.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnPrezziMSD.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnPrezziMSD.FlatAppearance.BorderSize = 0;
            this.btnPrezziMSD.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPrezziMSD.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPrezziMSD.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPrezziMSD.Location = new System.Drawing.Point(0, 68);
            this.btnPrezziMSD.Margin = new System.Windows.Forms.Padding(10);
            this.btnPrezziMSD.Name = "btnPrezziMSD";
            this.btnPrezziMSD.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnPrezziMSD.Size = new System.Drawing.Size(274, 34);
            this.btnPrezziMSD.TabIndex = 3;
            this.btnPrezziMSD.Text = "button3";
            this.btnPrezziMSD.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPrezziMSD.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnPrezziMSD.UseVisualStyleBackColor = true;
            // 
            // btnValidazioneTL
            // 
            this.btnValidazioneTL.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnValidazioneTL.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnValidazioneTL.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnValidazioneTL.FlatAppearance.BorderSize = 0;
            this.btnValidazioneTL.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnValidazioneTL.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnValidazioneTL.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnValidazioneTL.Location = new System.Drawing.Point(0, 102);
            this.btnValidazioneTL.Margin = new System.Windows.Forms.Padding(10);
            this.btnValidazioneTL.Name = "btnValidazioneTL";
            this.btnValidazioneTL.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnValidazioneTL.Size = new System.Drawing.Size(274, 34);
            this.btnValidazioneTL.TabIndex = 4;
            this.btnValidazioneTL.Text = "button4";
            this.btnValidazioneTL.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnValidazioneTL.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnValidazioneTL.UseVisualStyleBackColor = true;
            // 
            // btnPrevisioneCT
            // 
            this.btnPrevisioneCT.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnPrevisioneCT.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnPrevisioneCT.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnPrevisioneCT.FlatAppearance.BorderSize = 0;
            this.btnPrevisioneCT.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnPrevisioneCT.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnPrevisioneCT.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPrevisioneCT.Location = new System.Drawing.Point(0, 136);
            this.btnPrevisioneCT.Margin = new System.Windows.Forms.Padding(10);
            this.btnPrevisioneCT.Name = "btnPrevisioneCT";
            this.btnPrevisioneCT.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnPrevisioneCT.Size = new System.Drawing.Size(274, 34);
            this.btnPrevisioneCT.TabIndex = 5;
            this.btnPrevisioneCT.Text = "button5";
            this.btnPrevisioneCT.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnPrevisioneCT.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnPrevisioneCT.UseVisualStyleBackColor = true;
            // 
            // btnProgrImp
            // 
            this.btnProgrImp.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnProgrImp.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnProgrImp.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnProgrImp.FlatAppearance.BorderSize = 0;
            this.btnProgrImp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnProgrImp.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnProgrImp.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnProgrImp.Location = new System.Drawing.Point(0, 170);
            this.btnProgrImp.Margin = new System.Windows.Forms.Padding(10);
            this.btnProgrImp.Name = "btnProgrImp";
            this.btnProgrImp.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnProgrImp.Size = new System.Drawing.Size(274, 34);
            this.btnProgrImp.TabIndex = 6;
            this.btnProgrImp.Text = "button6";
            this.btnProgrImp.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnProgrImp.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnProgrImp.UseVisualStyleBackColor = true;
            // 
            // btnOfferteMGP
            // 
            this.btnOfferteMGP.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnOfferteMGP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnOfferteMGP.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnOfferteMGP.FlatAppearance.BorderSize = 0;
            this.btnOfferteMGP.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOfferteMGP.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOfferteMGP.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOfferteMGP.Location = new System.Drawing.Point(0, 204);
            this.btnOfferteMGP.Margin = new System.Windows.Forms.Padding(10);
            this.btnOfferteMGP.Name = "btnOfferteMGP";
            this.btnOfferteMGP.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnOfferteMGP.Size = new System.Drawing.Size(274, 34);
            this.btnOfferteMGP.TabIndex = 7;
            this.btnOfferteMGP.Text = "button7";
            this.btnOfferteMGP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOfferteMGP.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnOfferteMGP.UseVisualStyleBackColor = true;
            // 
            // btnOfferteMSD
            // 
            this.btnOfferteMSD.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnOfferteMSD.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnOfferteMSD.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnOfferteMSD.FlatAppearance.BorderSize = 0;
            this.btnOfferteMSD.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOfferteMSD.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOfferteMSD.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOfferteMSD.Location = new System.Drawing.Point(0, 238);
            this.btnOfferteMSD.Margin = new System.Windows.Forms.Padding(10);
            this.btnOfferteMSD.Name = "btnOfferteMSD";
            this.btnOfferteMSD.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnOfferteMSD.Size = new System.Drawing.Size(274, 34);
            this.btnOfferteMSD.TabIndex = 8;
            this.btnOfferteMSD.Text = "button8";
            this.btnOfferteMSD.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOfferteMSD.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnOfferteMSD.UseVisualStyleBackColor = true;
            // 
            // btnOfferteMB
            // 
            this.btnOfferteMB.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnOfferteMB.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnOfferteMB.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnOfferteMB.FlatAppearance.BorderSize = 0;
            this.btnOfferteMB.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnOfferteMB.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOfferteMB.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOfferteMB.Location = new System.Drawing.Point(0, 272);
            this.btnOfferteMB.Margin = new System.Windows.Forms.Padding(10);
            this.btnOfferteMB.Name = "btnOfferteMB";
            this.btnOfferteMB.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnOfferteMB.Size = new System.Drawing.Size(274, 34);
            this.btnOfferteMB.TabIndex = 9;
            this.btnOfferteMB.Text = "button9";
            this.btnOfferteMB.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnOfferteMB.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnOfferteMB.UseVisualStyleBackColor = true;
            // 
            // btnInvioProgrammi
            // 
            this.btnInvioProgrammi.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnInvioProgrammi.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnInvioProgrammi.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnInvioProgrammi.FlatAppearance.BorderSize = 0;
            this.btnInvioProgrammi.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnInvioProgrammi.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInvioProgrammi.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnInvioProgrammi.Location = new System.Drawing.Point(0, 306);
            this.btnInvioProgrammi.Margin = new System.Windows.Forms.Padding(10);
            this.btnInvioProgrammi.Name = "btnInvioProgrammi";
            this.btnInvioProgrammi.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnInvioProgrammi.Size = new System.Drawing.Size(274, 34);
            this.btnInvioProgrammi.TabIndex = 10;
            this.btnInvioProgrammi.Text = "button10";
            this.btnInvioProgrammi.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnInvioProgrammi.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnInvioProgrammi.UseVisualStyleBackColor = true;
            // 
            // btnSisCom
            // 
            this.btnSisCom.BackgroundImage = global::PSOLauncher.Properties.Resources.gas_icon;
            this.btnSisCom.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.btnSisCom.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnSisCom.FlatAppearance.BorderSize = 0;
            this.btnSisCom.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSisCom.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSisCom.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSisCom.Location = new System.Drawing.Point(0, 340);
            this.btnSisCom.Margin = new System.Windows.Forms.Padding(10);
            this.btnSisCom.Name = "btnSisCom";
            this.btnSisCom.Padding = new System.Windows.Forms.Padding(30, 0, 0, 0);
            this.btnSisCom.Size = new System.Drawing.Size(274, 34);
            this.btnSisCom.TabIndex = 11;
            this.btnSisCom.Text = "button11";
            this.btnSisCom.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnSisCom.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnSisCom.UseVisualStyleBackColor = true;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(274, 382);
            this.Controls.Add(this.btnSisCom);
            this.Controls.Add(this.btnInvioProgrammi);
            this.Controls.Add(this.btnOfferteMB);
            this.Controls.Add(this.btnOfferteMSD);
            this.Controls.Add(this.btnOfferteMGP);
            this.Controls.Add(this.btnProgrImp);
            this.Controls.Add(this.btnPrevisioneCT);
            this.Controls.Add(this.btnValidazioneTL);
            this.Controls.Add(this.btnPrezziMSD);
            this.Controls.Add(this.btnUnitComm);
            this.Controls.Add(this.btnPrevisioneGas);
            this.Name = "Form1";
            this.Opacity = 0D;
            this.ShowInTaskbar = false;
            this.Text = "PSO - Avvio programmi";
            this.menuTrayIcon.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.NotifyIcon IconTray;
        private System.Windows.Forms.ContextMenuStrip menuTrayIcon;
        private System.Windows.Forms.ToolStripMenuItem previsioneGas;
        private System.Windows.Forms.ToolStripMenuItem unitComm;
        private System.Windows.Forms.ToolStripMenuItem prezziMSD;
        private System.Windows.Forms.ToolStripMenuItem validazioneTL;
        private System.Windows.Forms.ToolStripMenuItem previsioneCT;
        private System.Windows.Forms.ToolStripMenuItem progrImp;
        private System.Windows.Forms.ToolStripMenuItem offerteMGP;
        private System.Windows.Forms.ToolStripMenuItem offerteMSD;
        private System.Windows.Forms.ToolStripMenuItem offerteMB;
        private System.Windows.Forms.ToolStripMenuItem invioProgrammi;
        private System.Windows.Forms.ToolStripMenuItem sisCom;
        private System.Windows.Forms.Button btnPrevisioneGas;
        private System.Windows.Forms.Button btnUnitComm;
        private System.Windows.Forms.Button btnPrezziMSD;
        private System.Windows.Forms.Button btnValidazioneTL;
        private System.Windows.Forms.Button btnPrevisioneCT;
        private System.Windows.Forms.Button btnProgrImp;
        private System.Windows.Forms.Button btnOfferteMGP;
        private System.Windows.Forms.Button btnOfferteMSD;
        private System.Windows.Forms.Button btnOfferteMB;
        private System.Windows.Forms.Button btnInvioProgrammi;
        private System.Windows.Forms.Button btnSisCom;

    }
}

