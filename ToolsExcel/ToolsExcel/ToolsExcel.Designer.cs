namespace Iren.FrontOffice.Tools
{
    partial class ToolsExcel : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ToolsExcel()
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
            this.TabFormulas = this.Factory.CreateRibbonTab();
            this.TabReview = this.Factory.CreateRibbonTab();
            this.TabDeveloper = this.Factory.CreateRibbonTab();
            this.TabPrintPreview = this.Factory.CreateRibbonTab();
            this.TabBackgroundRemoval = this.Factory.CreateRibbonTab();
            this.FrontOffice = this.Factory.CreateRibbonTab();
            this.TabFormulas.SuspendLayout();
            this.TabReview.SuspendLayout();
            this.TabDeveloper.SuspendLayout();
            this.TabPrintPreview.SuspendLayout();
            this.TabBackgroundRemoval.SuspendLayout();
            this.FrontOffice.SuspendLayout();
            // 
            // TabFormulas
            // 
            this.TabFormulas.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabFormulas.ControlId.OfficeId = "TabFormulas";
            this.TabFormulas.Name = "TabFormulas";
            // 
            // TabReview
            // 
            this.TabReview.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabReview.ControlId.OfficeId = "TabReview";
            this.TabReview.Name = "TabReview";
            // 
            // TabDeveloper
            // 
            this.TabDeveloper.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabDeveloper.ControlId.OfficeId = "TabDeveloper";
            this.TabDeveloper.Name = "TabDeveloper";
            // 
            // TabPrintPreview
            // 
            this.TabPrintPreview.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabPrintPreview.ControlId.OfficeId = "TabPrintPreview";
            this.TabPrintPreview.Label = "TabPrintPreview";
            this.TabPrintPreview.Name = "TabPrintPreview";
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
            this.FrontOffice.Label = "Front Office";
            this.FrontOffice.Name = "FrontOffice";
            // 
            // ToolsExcel
            // 
            this.Name = "ToolsExcel";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.StartFromScratch = true;
            this.Tabs.Add(this.TabFormulas);
            this.Tabs.Add(this.TabReview);
            this.Tabs.Add(this.TabDeveloper);
            this.Tabs.Add(this.TabPrintPreview);
            this.Tabs.Add(this.TabBackgroundRemoval);
            this.Tabs.Add(this.FrontOffice);
            this.TabFormulas.ResumeLayout(false);
            this.TabFormulas.PerformLayout();
            this.TabReview.ResumeLayout(false);
            this.TabReview.PerformLayout();
            this.TabDeveloper.ResumeLayout(false);
            this.TabDeveloper.PerformLayout();
            this.TabPrintPreview.ResumeLayout(false);
            this.TabPrintPreview.PerformLayout();
            this.TabBackgroundRemoval.ResumeLayout(false);
            this.TabBackgroundRemoval.PerformLayout();
            this.FrontOffice.ResumeLayout(false);
            this.FrontOffice.PerformLayout();

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab TabFormulas;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabReview;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabDeveloper;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabPrintPreview;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabBackgroundRemoval;
        private Microsoft.Office.Tools.Ribbon.RibbonTab FrontOffice;
    }

    partial class ThisRibbonCollection
    {
        internal ToolsExcel ToolsExcel
        {
            get { return this.GetRibbon<ToolsExcel>(); }
        }
    }
}
