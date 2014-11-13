namespace RiMoST2
{
    partial class RiMoST : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RiMoST()
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
            this.tbRichiestaModifica = this.Factory.CreateRibbonTab();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnChiudi = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnReset = this.Factory.CreateRibbonButton();
            this.btnInvia = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.tbRichiestaModifica.SuspendLayout();
            this.group2.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tbRichiestaModifica
            // 
            this.tbRichiestaModifica.Groups.Add(this.group2);
            this.tbRichiestaModifica.Groups.Add(this.group1);
            this.tbRichiestaModifica.Groups.Add(this.group3);
            this.tbRichiestaModifica.Label = "Richiesta Modifica";
            this.tbRichiestaModifica.Name = "tbRichiestaModifica";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnChiudi);
            this.group2.Label = "Chiudi";
            this.group2.Name = "group2";
            // 
            // btnChiudi
            // 
            this.btnChiudi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnChiudi.Image = global::RiMoST2.Properties.Resources.Close_icon;
            this.btnChiudi.Label = "Chiudi senza inviare";
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.ShowImage = true;
            this.btnChiudi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChiudi_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnReset);
            this.group1.Items.Add(this.btnInvia);
            this.group1.Label = "Modifica";
            this.group1.Name = "group1";
            // 
            // btnReset
            // 
            this.btnReset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReset.Image = global::RiMoST2.Properties.Resources.Reset_icon;
            this.btnReset.Label = "Cancella il contenuto del foglio";
            this.btnReset.Name = "btnReset";
            this.btnReset.ShowImage = true;
            this.btnReset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReset_Click);
            // 
            // btnInvia
            // 
            this.btnInvia.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInvia.Image = global::RiMoST2.Properties.Resources.Send_icon;
            this.btnInvia.Label = "Conferma e invia";
            this.btnInvia.Name = "btnInvia";
            this.btnInvia.ShowImage = true;
            this.btnInvia.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInvia_Click);
            // 
            // group3
            // 
            this.group3.Label = "group3";
            this.group3.Name = "group3";
            // 
            // RiMoST
            // 
            this.Name = "RiMoST";
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

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tbRichiestaModifica;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInvia;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReset;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChiudi;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
    }

    partial class ThisRibbonCollection
    {
        internal RiMoST RiMoST
        {
            get { return this.GetRibbon<RiMoST>(); }
        }
    }
}
