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
            this.group2 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnChiudi = this.Factory.CreateRibbonButton();
            this.btnInvia = this.Factory.CreateRibbonButton();
            this.btnReset = this.Factory.CreateRibbonButton();
            this.btnRefresh = this.Factory.CreateRibbonButton();
            this.btnPrint = this.Factory.CreateRibbonButton();
            this.btnAnnulla = this.Factory.CreateRibbonButton();
            this.tbRichiestaModifica.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            this.group4.SuspendLayout();
            // 
            // tbRichiestaModifica
            // 
            this.tbRichiestaModifica.Groups.Add(this.group2);
            this.tbRichiestaModifica.Groups.Add(this.group1);
            this.tbRichiestaModifica.Groups.Add(this.group4);
            this.tbRichiestaModifica.Label = "Richiesta Modifica";
            this.tbRichiestaModifica.Name = "tbRichiestaModifica";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnChiudi);
            this.group2.Items.Add(this.separator1);
            this.group2.Items.Add(this.btnInvia);
            this.group2.Label = "Azioni";
            this.group2.Name = "group2";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnReset);
            this.group1.Items.Add(this.btnRefresh);
            this.group1.Items.Add(this.btnPrint);
            this.group1.Label = "Modifica";
            this.group1.Name = "group1";
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnAnnulla);
            this.group4.Label = "Gestione Richieste";
            this.group4.Name = "group4";
            // 
            // btnChiudi
            // 
            this.btnChiudi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnChiudi.Image = ((System.Drawing.Image)(resources.GetObject("btnChiudi.Image")));
            this.btnChiudi.Label = "Chiudi";
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.ShowImage = true;
            this.btnChiudi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChiudi_Click);
            // 
            // btnInvia
            // 
            this.btnInvia.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInvia.Image = ((System.Drawing.Image)(resources.GetObject("btnInvia.Image")));
            this.btnInvia.Label = "Conferma e invia";
            this.btnInvia.Name = "btnInvia";
            this.btnInvia.ShowImage = true;
            this.btnInvia.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInvia_Click);
            // 
            // btnReset
            // 
            this.btnReset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReset.Image = global::RiMoST2.Properties.Resources.Eraser_icon;
            this.btnReset.Label = "Cancella Modifiche";
            this.btnReset.Name = "btnReset";
            this.btnReset.ShowImage = true;
            this.btnReset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReset_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRefresh.Image = ((System.Drawing.Image)(resources.GetObject("btnRefresh.Image")));
            this.btnRefresh.Label = "Aggiorna n°";
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.ShowImage = true;
            this.btnRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRefresh_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPrint.Image = ((System.Drawing.Image)(resources.GetObject("btnPrint.Image")));
            this.btnPrint.Label = "Stampa";
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.ShowImage = true;
            this.btnPrint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrint_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAnnulla.Image = global::RiMoST2.Properties.Resources.Bin_icon;
            this.btnAnnulla.Label = "Annulla una richiesta";
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.ShowImage = true;
            this.btnAnnulla.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAnnulla_Click);
            // 
            // RiMoST_Ribbon
            // 
            this.Name = "RiMoST_Ribbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.StartFromScratch = true;
            this.Tabs.Add(this.tbRichiestaModifica);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RiMoST_Load);
            this.tbRichiestaModifica.ResumeLayout(false);
            this.tbRichiestaModifica.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();

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
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal RiMoST_Ribbon RiMoST
        {
            get { return this.GetRibbon<RiMoST_Ribbon>(); }
        }
    }
}
