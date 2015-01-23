﻿namespace Iren.FrontOffice.Tools
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btnAggiornaStruttura = this.Factory.CreateRibbonButton();
            this.btnCalendar = this.Factory.CreateRibbonButton();
            this.btnRampe = this.Factory.CreateRibbonButton();
            this.btnAggiornaDati = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnChiudi = this.Factory.CreateRibbonButton();
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
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
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
            this.FrontOffice.Groups.Add(this.group3);
            this.FrontOffice.Groups.Add(this.groupAggiorna);
            this.FrontOffice.Groups.Add(this.group1);
            this.FrontOffice.Groups.Add(this.group2);
            this.FrontOffice.Label = "Front Office";
            this.FrontOffice.Name = "FrontOffice";
            // 
            // groupAggiorna
            // 
            this.groupAggiorna.Items.Add(this.btnAggiornaDati);
            this.groupAggiorna.Items.Add(this.btnAggiornaStruttura);
            this.groupAggiorna.Label = "Aggiorna";
            this.groupAggiorna.Name = "groupAggiorna";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnCalendar);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnRampe);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
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
            // btnCalendar
            // 
            this.btnCalendar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCalendar.Description = "Apre il calendario per cambiare la data";
            this.btnCalendar.Image = global::ToolsExcel.Properties.Resources.Calendar_icon;
            this.btnCalendar.Label = "Calendario";
            this.btnCalendar.Name = "btnCalendar";
            this.btnCalendar.ScreenTip = "Apre il calendario per cambiare la data";
            this.btnCalendar.ShowImage = true;
            this.btnCalendar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalendar_Click);
            // 
            // btnRampe
            // 
            this.btnRampe.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRampe.Image = global::ToolsExcel.Properties.Resources.checklist_icon;
            this.btnRampe.Label = "Seleziona Rampa";
            this.btnRampe.Name = "btnRampe";
            this.btnRampe.ShowImage = true;
            this.btnRampe.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRampe_Click);
            // 
            // btnAggiornaDati
            // 
            this.btnAggiornaDati.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAggiornaDati.Image = global::ToolsExcel.Properties.Resources.Generate_tables_icon;
            this.btnAggiornaDati.Label = "Aggiorna Dati";
            this.btnAggiornaDati.Name = "btnAggiornaDati";
            this.btnAggiornaDati.ShowImage = true;
            this.btnAggiornaDati.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAggiornaDati_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnChiudi);
            this.group3.Label = "group3";
            this.group3.Name = "group3";
            // 
            // btnChiudi
            // 
            this.btnChiudi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnChiudi.Image = global::ToolsExcel.Properties.Resources.Save_icon;
            this.btnChiudi.Label = "Chiudi";
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.ShowImage = true;
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
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();

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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalendar;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRampe;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChiudi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAggiornaDati;
    }

    partial class ThisRibbonCollection
    {
        internal ToolsExcelRibbon ToolsExcelRibbon
        {
            get { return this.GetRibbon<ToolsExcelRibbon>(); }
        }
    }
}
