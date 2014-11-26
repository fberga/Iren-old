namespace Iren.FrontOffice.Tools
{
    partial class ToolsExcelRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ToolsExcelRibbon()
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
            this.TabHome = this.Factory.CreateRibbonTab();
            this.TabInsert = this.Factory.CreateRibbonTab();
            this.TabPageLayoutExcel = this.Factory.CreateRibbonTab();
            this.TabFormulas = this.Factory.CreateRibbonTab();
            this.TabData = this.Factory.CreateRibbonTab();
            this.TabReview = this.Factory.CreateRibbonTab();
            this.TabView = this.Factory.CreateRibbonTab();
            this.TabDeveloper = this.Factory.CreateRibbonTab();
            this.TabAddIns = this.Factory.CreateRibbonTab();
            this.TabPrintPreview = this.Factory.CreateRibbonTab();
            this.TabBackgroundRemoval = this.Factory.CreateRibbonTab();
            this.FrontOffice = this.Factory.CreateRibbonTab();
            this.groupAggiorna = this.Factory.CreateRibbonGroup();
            this.btnAggiornaStruttura = this.Factory.CreateRibbonButton();
            this.TabHome.SuspendLayout();
            this.TabInsert.SuspendLayout();
            this.TabPageLayoutExcel.SuspendLayout();
            this.TabFormulas.SuspendLayout();
            this.TabData.SuspendLayout();
            this.TabReview.SuspendLayout();
            this.TabView.SuspendLayout();
            this.TabDeveloper.SuspendLayout();
            this.TabAddIns.SuspendLayout();
            this.TabPrintPreview.SuspendLayout();
            this.TabBackgroundRemoval.SuspendLayout();
            this.FrontOffice.SuspendLayout();
            this.groupAggiorna.SuspendLayout();
            // 
            // TabHome
            // 
            this.TabHome.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabHome.ControlId.OfficeId = "TabHome";
            this.TabHome.Label = "TabHome";
            this.TabHome.Name = "TabHome";
            this.TabHome.Visible = false;
            // 
            // TabInsert
            // 
            this.TabInsert.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabInsert.ControlId.OfficeId = "TabInsert";
            this.TabInsert.Label = "TabInsert";
            this.TabInsert.Name = "TabInsert";
            this.TabInsert.Visible = false;
            // 
            // TabPageLayoutExcel
            // 
            this.TabPageLayoutExcel.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabPageLayoutExcel.ControlId.OfficeId = "TabPageLayoutExcel";
            this.TabPageLayoutExcel.Label = "TabPageLayoutExcel";
            this.TabPageLayoutExcel.Name = "TabPageLayoutExcel";
            this.TabPageLayoutExcel.Visible = false;
            // 
            // TabFormulas
            // 
            this.TabFormulas.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabFormulas.ControlId.OfficeId = "TabFormulas";
            this.TabFormulas.Label = "TabFormulas";
            this.TabFormulas.Name = "TabFormulas";
            // 
            // TabData
            // 
            this.TabData.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabData.ControlId.OfficeId = "TabData";
            this.TabData.Label = "TabData";
            this.TabData.Name = "TabData";
            this.TabData.Visible = false;
            // 
            // TabReview
            // 
            this.TabReview.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabReview.ControlId.OfficeId = "TabReview";
            this.TabReview.Label = "TabReview";
            this.TabReview.Name = "TabReview";
            // 
            // TabView
            // 
            this.TabView.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabView.ControlId.OfficeId = "TabView";
            this.TabView.Label = "TabView";
            this.TabView.Name = "TabView";
            this.TabView.Visible = false;
            // 
            // TabDeveloper
            // 
            this.TabDeveloper.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabDeveloper.ControlId.OfficeId = "TabDeveloper";
            this.TabDeveloper.Label = "TabDeveloper";
            this.TabDeveloper.Name = "TabDeveloper";
            // 
            // TabAddIns
            // 
            this.TabAddIns.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabAddIns.Label = "TabAddIns";
            this.TabAddIns.Name = "TabAddIns";
            this.TabAddIns.Visible = false;
            // 
            // TabPrintPreview
            // 
            this.TabPrintPreview.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabPrintPreview.ControlId.OfficeId = "TabPrintPreview";
            this.TabPrintPreview.Label = "TabPrintPreview";
            this.TabPrintPreview.Name = "TabPrintPreview";
            this.TabPrintPreview.Visible = false;
            // 
            // TabBackgroundRemoval
            // 
            this.TabBackgroundRemoval.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabBackgroundRemoval.ControlId.OfficeId = "TabBackgroundRemoval";
            this.TabBackgroundRemoval.Label = "TabBackgroundRemoval";
            this.TabBackgroundRemoval.Name = "TabBackgroundRemoval";
            this.TabBackgroundRemoval.Visible = false;
            // 
            // FrontOffice
            // 
            this.FrontOffice.Groups.Add(this.groupAggiorna);
            this.FrontOffice.Label = "Front Office";
            this.FrontOffice.Name = "FrontOffice";
            // 
            // groupAggiorna
            // 
            this.groupAggiorna.Items.Add(this.btnAggiornaStruttura);
            this.groupAggiorna.Label = "Aggiorna";
            this.groupAggiorna.Name = "groupAggiorna";
            // 
            // btnAggiornaStruttura
            // 
            this.btnAggiornaStruttura.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAggiornaStruttura.Image = global::ToolsExcel.Properties.Resources.Structure_Refresh_icon;
            this.btnAggiornaStruttura.Label = "Aggiorna Struttura";
            this.btnAggiornaStruttura.Name = "btnAggiornaStruttura";
            this.btnAggiornaStruttura.ShowImage = true;
            this.btnAggiornaStruttura.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAggiornaStruttura_Click);
            // 
            // ToolsExcelRibbon
            // 
            this.Name = "ToolsExcelRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.StartFromScratch = true;
            this.Tabs.Add(this.FrontOffice);
            this.Tabs.Add(this.TabHome);
            this.Tabs.Add(this.TabInsert);
            this.Tabs.Add(this.TabPageLayoutExcel);
            this.Tabs.Add(this.TabFormulas);
            this.Tabs.Add(this.TabData);
            this.Tabs.Add(this.TabReview);
            this.Tabs.Add(this.TabView);
            this.Tabs.Add(this.TabDeveloper);
            this.Tabs.Add(this.TabAddIns);
            this.Tabs.Add(this.TabPrintPreview);
            this.Tabs.Add(this.TabBackgroundRemoval);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ToolsExcelRibbon_Load);
            this.TabHome.ResumeLayout(false);
            this.TabHome.PerformLayout();
            this.TabInsert.ResumeLayout(false);
            this.TabInsert.PerformLayout();
            this.TabPageLayoutExcel.ResumeLayout(false);
            this.TabPageLayoutExcel.PerformLayout();
            this.TabFormulas.ResumeLayout(false);
            this.TabFormulas.PerformLayout();
            this.TabData.ResumeLayout(false);
            this.TabData.PerformLayout();
            this.TabReview.ResumeLayout(false);
            this.TabReview.PerformLayout();
            this.TabView.ResumeLayout(false);
            this.TabView.PerformLayout();
            this.TabDeveloper.ResumeLayout(false);
            this.TabDeveloper.PerformLayout();
            this.TabAddIns.ResumeLayout(false);
            this.TabAddIns.PerformLayout();
            this.TabPrintPreview.ResumeLayout(false);
            this.TabPrintPreview.PerformLayout();
            this.TabBackgroundRemoval.ResumeLayout(false);
            this.TabBackgroundRemoval.PerformLayout();
            this.FrontOffice.ResumeLayout(false);
            this.FrontOffice.PerformLayout();
            this.groupAggiorna.ResumeLayout(false);
            this.groupAggiorna.PerformLayout();

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab TabHome;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabInsert;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabPageLayoutExcel;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabFormulas;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabData;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabReview;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabView;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabDeveloper;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabAddIns;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabPrintPreview;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabBackgroundRemoval;
        private Microsoft.Office.Tools.Ribbon.RibbonTab FrontOffice;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAggiorna;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAggiornaStruttura;
    }

    partial class ThisRibbonCollection
    {
        internal ToolsExcelRibbon ToolsExcelRibbon
        {
            get { return this.GetRibbon<ToolsExcelRibbon>(); }
        }
    }
}
