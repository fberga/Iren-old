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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnApri = this.Factory.CreateRibbonButton();
            this.btnSalva = this.Factory.CreateRibbonButton();
            this.btnReset = this.Factory.CreateRibbonButton();
            this.btnInvia = this.Factory.CreateRibbonButton();
            this.tbRichiestaModifica.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tbRichiestaModifica
            // 
            this.tbRichiestaModifica.Groups.Add(this.group1);
            this.tbRichiestaModifica.Label = "Richiesta Modifica";
            this.tbRichiestaModifica.Name = "tbRichiestaModifica";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnApri);
            this.group1.Items.Add(this.btnSalva);
            this.group1.Items.Add(this.btnReset);
            this.group1.Items.Add(this.btnInvia);
            this.group1.Label = "Modifica";
            this.group1.Name = "group1";
            // 
            // btnApri
            // 
            this.btnApri.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnApri.Image = global::RiMoST2.Properties.Resources.open_icon;
            this.btnApri.Label = "Apri";
            this.btnApri.Name = "btnApri";
            this.btnApri.ShowImage = true;
            // 
            // btnSalva
            // 
            this.btnSalva.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSalva.Image = global::RiMoST2.Properties.Resources.save_icon;
            this.btnSalva.Label = "Salva su questo computer";
            this.btnSalva.Name = "btnSalva";
            this.btnSalva.ShowImage = true;
            // 
            // btnReset
            // 
            this.btnReset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReset.Image = global::RiMoST2.Properties.Resources.reset_icon;
            this.btnReset.Label = "Cancella il contenuto del foglio";
            this.btnReset.Name = "btnReset";
            this.btnReset.ShowImage = true;
            // 
            // btnInvia
            // 
            this.btnInvia.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInvia.Image = global::RiMoST2.Properties.Resources.send_icon;
            this.btnInvia.Label = "Conferma e invia";
            this.btnInvia.Name = "btnInvia";
            this.btnInvia.ShowImage = true;
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
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tbRichiestaModifica;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInvia;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnApri;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReset;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSalva;
    }

    partial class ThisRibbonCollection
    {
        internal RiMoST RiMoST
        {
            get { return this.GetRibbon<RiMoST>(); }
        }
    }
}
