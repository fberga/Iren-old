namespace RiMoST2
{
    partial class RiMoST_Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RiMoST_Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Codice generato da Progettazione componenti

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RiMoST_Ribbon));
            this.tbRichiestaModifica = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.cbAnniDisponibili = this.Factory.CreateRibbonComboBox();
            this.chkIsDraft = this.Factory.CreateRibbonCheckBox();
            this.tabVersione = this.Factory.CreateRibbonTab();
            this.groupVersione = this.Factory.CreateRibbonGroup();
            this.lbVersioneLabel = this.Factory.CreateRibbonLabel();
            this.lbVersioneApp = this.Factory.CreateRibbonLabel();
            this.lbCoreV = this.Factory.CreateRibbonLabel();
            this.btnChiudi = this.Factory.CreateRibbonButton();
            this.btnInvia = this.Factory.CreateRibbonButton();
            this.btnSalvaBozza = this.Factory.CreateRibbonButton();
            this.btnReset = this.Factory.CreateRibbonButton();
            this.btnRefresh = this.Factory.CreateRibbonButton();
            this.btnPrint = this.Factory.CreateRibbonButton();
            this.btnAnnulla = this.Factory.CreateRibbonButton();
            this.btnModifica = this.Factory.CreateRibbonButton();
            this.tbRichiestaModifica.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group4.SuspendLayout();
            this.tabVersione.SuspendLayout();
            this.groupVersione.SuspendLayout();
            // 
            // tbRichiestaModifica
            // 
            this.tbRichiestaModifica.Groups.Add(this.group3);
            this.tbRichiestaModifica.Groups.Add(this.group2);
            this.tbRichiestaModifica.Groups.Add(this.group1);
            this.tbRichiestaModifica.Groups.Add(this.group4);
            resources.ApplyResources(this.tbRichiestaModifica, "tbRichiestaModifica");
            this.tbRichiestaModifica.Name = "tbRichiestaModifica";
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnChiudi);
            resources.ApplyResources(this.group3, "group3");
            this.group3.Name = "group3";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnInvia);
            this.group2.Items.Add(this.btnSalvaBozza);
            resources.ApplyResources(this.group2, "group2");
            this.group2.Name = "group2";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnReset);
            this.group1.Items.Add(this.btnRefresh);
            this.group1.Items.Add(this.btnPrint);
            resources.ApplyResources(this.group1, "group1");
            this.group1.Name = "group1";
            // 
            // group4
            // 
            this.group4.Items.Add(this.cbAnniDisponibili);
            this.group4.Items.Add(this.chkIsDraft);
            this.group4.Items.Add(this.btnAnnulla);
            this.group4.Items.Add(this.btnModifica);
            resources.ApplyResources(this.group4, "group4");
            this.group4.Name = "group4";
            // 
            // cbAnniDisponibili
            // 
            this.cbAnniDisponibili.Image = global::RiMoST2.Properties.Resources.calendar_icon;
            resources.ApplyResources(this.cbAnniDisponibili, "cbAnniDisponibili");
            this.cbAnniDisponibili.Name = "cbAnniDisponibili";
            this.cbAnniDisponibili.ShowImage = true;
            this.cbAnniDisponibili.ShowLabel = false;
            // 
            // chkIsDraft
            // 
            resources.ApplyResources(this.chkIsDraft, "chkIsDraft");
            this.chkIsDraft.Name = "chkIsDraft";
            // 
            // tabVersione
            // 
            this.tabVersione.Groups.Add(this.groupVersione);
            resources.ApplyResources(this.tabVersione, "tabVersione");
            this.tabVersione.Name = "tabVersione";
            // 
            // groupVersione
            // 
            this.groupVersione.Items.Add(this.lbVersioneLabel);
            this.groupVersione.Items.Add(this.lbVersioneApp);
            this.groupVersione.Items.Add(this.lbCoreV);
            resources.ApplyResources(this.groupVersione, "groupVersione");
            this.groupVersione.Name = "groupVersione";
            // 
            // lbVersioneLabel
            // 
            resources.ApplyResources(this.lbVersioneLabel, "lbVersioneLabel");
            this.lbVersioneLabel.Name = "lbVersioneLabel";
            // 
            // lbVersioneApp
            // 
            resources.ApplyResources(this.lbVersioneApp, "lbVersioneApp");
            this.lbVersioneApp.Name = "lbVersioneApp";
            // 
            // lbCoreV
            // 
            resources.ApplyResources(this.lbCoreV, "lbCoreV");
            this.lbCoreV.Name = "lbCoreV";
            // 
            // btnChiudi
            // 
            this.btnChiudi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.btnChiudi, "btnChiudi");
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.ShowImage = true;
            this.btnChiudi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChiudi_Click);
            // 
            // btnInvia
            // 
            this.btnInvia.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.btnInvia, "btnInvia");
            this.btnInvia.Name = "btnInvia";
            this.btnInvia.ShowImage = true;
            this.btnInvia.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInvia_Click);
            // 
            // btnSalvaBozza
            // 
            this.btnSalvaBozza.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSalvaBozza.Image = global::RiMoST2.Properties.Resources.save_icon;
            resources.ApplyResources(this.btnSalvaBozza, "btnSalvaBozza");
            this.btnSalvaBozza.Name = "btnSalvaBozza";
            this.btnSalvaBozza.ShowImage = true;
            this.btnSalvaBozza.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSalvaBozza_Click);
            // 
            // btnReset
            // 
            this.btnReset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReset.Image = global::RiMoST2.Properties.Resources.Eraser_icon;
            resources.ApplyResources(this.btnReset, "btnReset");
            this.btnReset.Name = "btnReset";
            this.btnReset.ShowImage = true;
            this.btnReset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReset_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.btnRefresh, "btnRefresh");
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.ShowImage = true;
            this.btnRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRefresh_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            resources.ApplyResources(this.btnPrint, "btnPrint");
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.ShowImage = true;
            this.btnPrint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrint_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAnnulla.Image = global::RiMoST2.Properties.Resources.Bin_icon;
            resources.ApplyResources(this.btnAnnulla, "btnAnnulla");
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.ShowImage = true;
            this.btnAnnulla.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAnnulla_Click);
            // 
            // btnModifica
            // 
            this.btnModifica.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnModifica.Image = global::RiMoST2.Properties.Resources.edit_icon;
            resources.ApplyResources(this.btnModifica, "btnModifica");
            this.btnModifica.Name = "btnModifica";
            this.btnModifica.ShowImage = true;
            this.btnModifica.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifica_Click);
            // 
            // RiMoST_Ribbon
            // 
            this.Name = "RiMoST_Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.StartFromScratch = true;
            this.Tabs.Add(this.tbRichiestaModifica);
            this.Tabs.Add(this.tabVersione);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RiMoST_Load);
            this.tbRichiestaModifica.ResumeLayout(false);
            this.tbRichiestaModifica.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.tabVersione.ResumeLayout(false);
            this.tabVersione.PerformLayout();
            this.groupVersione.ResumeLayout(false);
            this.groupVersione.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tbRichiestaModifica;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInvia;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReset;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChiudi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrint;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAnnulla;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifica;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSalvaBozza;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonCheckBox chkIsDraft;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cbAnniDisponibili;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabVersione;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupVersione;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbVersioneApp;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbVersioneLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbCoreV;
    }

    partial class ThisRibbonCollection
    {
        internal RiMoST_Ribbon RiMoST
        {
            get { return this.GetRibbon<RiMoST_Ribbon>(); }
        }
    }
}
