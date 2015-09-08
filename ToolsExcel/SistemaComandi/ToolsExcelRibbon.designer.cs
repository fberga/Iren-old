namespace Iren.ToolsExcel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ToolsExcelRibbon));
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            this.FrontOffice = this.Factory.CreateRibbonTab();
            this.groupChiudi = this.Factory.CreateRibbonGroup();
            this.btnChiudi = this.Factory.CreateRibbonButton();
            this.groupEmergenza = this.Factory.CreateRibbonGroup();
            this.btnForzaEmergenza = this.Factory.CreateRibbonToggleButton();
            this.btnEsportaXML = this.Factory.CreateRibbonButton();
            this.btnImportaXML = this.Factory.CreateRibbonButton();
            this.groupConfigura = this.Factory.CreateRibbonGroup();
            this.btnConfigura = this.Factory.CreateRibbonButton();
            this.btnConfiguraParametri = this.Factory.CreateRibbonButton();
            this.groupCalendario = this.Factory.CreateRibbonGroup();
            this.btnCalendar = this.Factory.CreateRibbonButton();
            this.groupModifica = this.Factory.CreateRibbonGroup();
            this.btnModifica = this.Factory.CreateRibbonToggleButton();
            this.groupAggiorna = this.Factory.CreateRibbonGroup();
            this.btnAggiornaDati = this.Factory.CreateRibbonButton();
            this.btnAggiornaStruttura = this.Factory.CreateRibbonButton();
            this.groupAzioni = this.Factory.CreateRibbonGroup();
            this.btnAzioni = this.Factory.CreateRibbonButton();
            this.btnOttimizza = this.Factory.CreateRibbonButton();
            this.btnRampe = this.Factory.CreateRibbonButton();
            this.groupAmbienti = this.Factory.CreateRibbonGroup();
            this.btnDev = this.Factory.CreateRibbonToggleButton();
            this.btnTest = this.Factory.CreateRibbonToggleButton();
            this.btnProduzione = this.Factory.CreateRibbonToggleButton();
            this.groupApplicativi = this.Factory.CreateRibbonGroup();
            this.btnPrevisioneGas = this.Factory.CreateRibbonToggleButton();
            this.btnUnitCommitment = this.Factory.CreateRibbonToggleButton();
            this.btnPrezziMSD = this.Factory.CreateRibbonToggleButton();
            this.btnValidazioneTL = this.Factory.CreateRibbonToggleButton();
            this.btnPrevisioneCT = this.Factory.CreateRibbonToggleButton();
            this.btnProgrammazioneImpianti = this.Factory.CreateRibbonToggleButton();
            this.btnOfferteMGP = this.Factory.CreateRibbonToggleButton();
            this.btnOfferteMSD = this.Factory.CreateRibbonToggleButton();
            this.btnOfferteMB = this.Factory.CreateRibbonToggleButton();
            this.btnInvioProgrammi = this.Factory.CreateRibbonToggleButton();
            this.btnSistemaComandi = this.Factory.CreateRibbonToggleButton();
            this.groupErrori = this.Factory.CreateRibbonGroup();
            this.btnMostraErrorPane = this.Factory.CreateRibbonButton();
            this.groupMSD = this.Factory.CreateRibbonGroup();
            this.labelMSD = this.Factory.CreateRibbonLabel();
            this.cmbMSD = this.Factory.CreateRibbonComboBox();
            this.groupStagione = this.Factory.CreateRibbonGroup();
            this.labelStagione = this.Factory.CreateRibbonLabel();
            this.cmbStagione = this.Factory.CreateRibbonComboBox();
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
            this.TabSmartArtToolsDesign = this.Factory.CreateRibbonTab();
            this.FrontOffice.SuspendLayout();
            this.groupChiudi.SuspendLayout();
            this.groupEmergenza.SuspendLayout();
            this.groupConfigura.SuspendLayout();
            this.groupCalendario.SuspendLayout();
            this.groupModifica.SuspendLayout();
            this.groupAggiorna.SuspendLayout();
            this.groupAzioni.SuspendLayout();
            this.groupAmbienti.SuspendLayout();
            this.groupApplicativi.SuspendLayout();
            this.groupErrori.SuspendLayout();
            this.groupMSD.SuspendLayout();
            this.groupStagione.SuspendLayout();
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
            this.TabSmartArtToolsDesign.SuspendLayout();
            // 
            // FrontOffice
            // 
            this.FrontOffice.Groups.Add(this.groupChiudi);
            this.FrontOffice.Groups.Add(this.groupEmergenza);
            this.FrontOffice.Groups.Add(this.groupConfigura);
            this.FrontOffice.Groups.Add(this.groupCalendario);
            this.FrontOffice.Groups.Add(this.groupModifica);
            this.FrontOffice.Groups.Add(this.groupAggiorna);
            this.FrontOffice.Groups.Add(this.groupAzioni);
            this.FrontOffice.Groups.Add(this.groupAmbienti);
            this.FrontOffice.Groups.Add(this.groupApplicativi);
            this.FrontOffice.Groups.Add(this.groupErrori);
            this.FrontOffice.Groups.Add(this.groupMSD);
            this.FrontOffice.Groups.Add(this.groupStagione);
            this.FrontOffice.Label = "Front Office";
            this.FrontOffice.Name = "FrontOffice";
            // 
            // groupChiudi
            // 
            this.groupChiudi.Items.Add(this.btnChiudi);
            this.groupChiudi.Label = " Chiudi";
            this.groupChiudi.Name = "groupChiudi";
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
            // groupEmergenza
            // 
            this.groupEmergenza.Items.Add(this.btnForzaEmergenza);
            this.groupEmergenza.Items.Add(this.btnEsportaXML);
            this.groupEmergenza.Items.Add(this.btnImportaXML);
            this.groupEmergenza.Label = "Emergenza";
            this.groupEmergenza.Name = "groupEmergenza";
            // 
            // btnForzaEmergenza
            // 
            this.btnForzaEmergenza.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnForzaEmergenza.Image = ((System.Drawing.Image)(resources.GetObject("btnForzaEmergenza.Image")));
            this.btnForzaEmergenza.Label = "Forza Emergenza";
            this.btnForzaEmergenza.Name = "btnForzaEmergenza";
            this.btnForzaEmergenza.ShowImage = true;
            this.btnForzaEmergenza.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnForzaEmergenza_Click);
            // 
            // btnEsportaXML
            // 
            this.btnEsportaXML.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnEsportaXML.Image = ((System.Drawing.Image)(resources.GetObject("btnEsportaXML.Image")));
            this.btnEsportaXML.Label = "Esporta dati in XML";
            this.btnEsportaXML.Name = "btnEsportaXML";
            this.btnEsportaXML.ShowImage = true;
            this.btnEsportaXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEsportaXML_Click);
            // 
            // btnImportaXML
            // 
            this.btnImportaXML.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnImportaXML.Image = ((System.Drawing.Image)(resources.GetObject("btnImportaXML.Image")));
            this.btnImportaXML.Label = "Importa dati da XML";
            this.btnImportaXML.Name = "btnImportaXML";
            this.btnImportaXML.ShowImage = true;
            this.btnImportaXML.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnImportaXML_Click);
            // 
            // groupConfigura
            // 
            this.groupConfigura.Items.Add(this.btnConfigura);
            this.groupConfigura.Items.Add(this.btnConfiguraParametri);
            this.groupConfigura.Label = "Configura";
            this.groupConfigura.Name = "groupConfigura";
            // 
            // btnConfigura
            // 
            this.btnConfigura.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConfigura.Image = ((System.Drawing.Image)(resources.GetObject("btnConfigura.Image")));
            this.btnConfigura.Label = "Configura percorsi";
            this.btnConfigura.Name = "btnConfigura";
            this.btnConfigura.ShowImage = true;
            this.btnConfigura.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConfigura_Click);
            // 
            // btnConfiguraParametri
            // 
            this.btnConfiguraParametri.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnConfiguraParametri.Image = ((System.Drawing.Image)(resources.GetObject("btnConfiguraParametri.Image")));
            this.btnConfiguraParametri.Label = "Configura parametri";
            this.btnConfiguraParametri.Name = "btnConfiguraParametri";
            this.btnConfiguraParametri.ShowImage = true;
            this.btnConfiguraParametri.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConfiguraParametri_Click);
            // 
            // groupCalendario
            // 
            this.groupCalendario.Items.Add(this.btnCalendar);
            this.groupCalendario.Label = "Calendario";
            this.groupCalendario.Name = "groupCalendario";
            // 
            // btnCalendar
            // 
            this.btnCalendar.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCalendar.Description = "Apre il calendario per cambiare la data";
            this.btnCalendar.Image = ((System.Drawing.Image)(resources.GetObject("btnCalendar.Image")));
            this.btnCalendar.Label = "Calendario";
            this.btnCalendar.Name = "btnCalendar";
            this.btnCalendar.ScreenTip = "Apre il calendario per cambiare la data";
            this.btnCalendar.ShowImage = true;
            this.btnCalendar.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCalendar_Click);
            // 
            // groupModifica
            // 
            this.groupModifica.Items.Add(this.btnModifica);
            this.groupModifica.Label = "Modifica";
            this.groupModifica.Name = "groupModifica";
            // 
            // btnModifica
            // 
            this.btnModifica.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnModifica.Image = ((System.Drawing.Image)(resources.GetObject("btnModifica.Image")));
            this.btnModifica.Label = "Modifica NO";
            this.btnModifica.Name = "btnModifica";
            this.btnModifica.ShowImage = true;
            this.btnModifica.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifica_Click);
            // 
            // groupAggiorna
            // 
            this.groupAggiorna.Items.Add(this.btnAggiornaDati);
            this.groupAggiorna.Items.Add(this.btnAggiornaStruttura);
            this.groupAggiorna.Label = "Aggiorna";
            this.groupAggiorna.Name = "groupAggiorna";
            // 
            // btnAggiornaDati
            // 
            this.btnAggiornaDati.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAggiornaDati.Image = ((System.Drawing.Image)(resources.GetObject("btnAggiornaDati.Image")));
            this.btnAggiornaDati.Label = "Aggiorna Dati";
            this.btnAggiornaDati.Name = "btnAggiornaDati";
            this.btnAggiornaDati.ShowImage = true;
            this.btnAggiornaDati.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAggiornaDati_Click);
            // 
            // btnAggiornaStruttura
            // 
            this.btnAggiornaStruttura.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAggiornaStruttura.Image = ((System.Drawing.Image)(resources.GetObject("btnAggiornaStruttura.Image")));
            this.btnAggiornaStruttura.Label = "Aggiorna Struttura";
            this.btnAggiornaStruttura.Name = "btnAggiornaStruttura";
            this.btnAggiornaStruttura.ShowImage = true;
            this.btnAggiornaStruttura.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAggiornaStruttura_Click);
            // 
            // groupAzioni
            // 
            this.groupAzioni.Items.Add(this.btnAzioni);
            this.groupAzioni.Items.Add(this.btnOttimizza);
            this.groupAzioni.Items.Add(this.btnRampe);
            this.groupAzioni.Label = "Azioni";
            this.groupAzioni.Name = "groupAzioni";
            // 
            // btnAzioni
            // 
            this.btnAzioni.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAzioni.Image = ((System.Drawing.Image)(resources.GetObject("btnAzioni.Image")));
            this.btnAzioni.Label = "Start";
            this.btnAzioni.Name = "btnAzioni";
            this.btnAzioni.ShowImage = true;
            this.btnAzioni.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAzioni_Click);
            // 
            // btnOttimizza
            // 
            this.btnOttimizza.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOttimizza.Image = ((System.Drawing.Image)(resources.GetObject("btnOttimizza.Image")));
            this.btnOttimizza.Label = "Ottimizza";
            this.btnOttimizza.Name = "btnOttimizza";
            this.btnOttimizza.ShowImage = true;
            this.btnOttimizza.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOttimizza_Click);
            // 
            // btnRampe
            // 
            this.btnRampe.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRampe.Image = ((System.Drawing.Image)(resources.GetObject("btnRampe.Image")));
            this.btnRampe.Label = "Rampe";
            this.btnRampe.Name = "btnRampe";
            this.btnRampe.ShowImage = true;
            this.btnRampe.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRampe_Click);
            // 
            // groupAmbienti
            // 
            this.groupAmbienti.Items.Add(this.btnDev);
            this.groupAmbienti.Items.Add(this.btnTest);
            this.groupAmbienti.Items.Add(this.btnProduzione);
            this.groupAmbienti.Label = "Ambienti";
            this.groupAmbienti.Name = "groupAmbienti";
            // 
            // btnDev
            // 
            this.btnDev.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnDev.Image = ((System.Drawing.Image)(resources.GetObject("btnDev.Image")));
            this.btnDev.Label = "Dev";
            this.btnDev.Name = "btnDev";
            this.btnDev.ShowImage = true;
            this.btnDev.Visible = false;
            this.btnDev.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelezionaAmbiente_Click);
            // 
            // btnTest
            // 
            this.btnTest.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnTest.Image = ((System.Drawing.Image)(resources.GetObject("btnTest.Image")));
            this.btnTest.Label = "Test";
            this.btnTest.Name = "btnTest";
            this.btnTest.ShowImage = true;
            this.btnTest.Visible = false;
            this.btnTest.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelezionaAmbiente_Click);
            // 
            // btnProduzione
            // 
            this.btnProduzione.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnProduzione.Image = ((System.Drawing.Image)(resources.GetObject("btnProduzione.Image")));
            this.btnProduzione.Label = "Prod";
            this.btnProduzione.Name = "btnProduzione";
            this.btnProduzione.ShowImage = true;
            this.btnProduzione.Visible = false;
            this.btnProduzione.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSelezionaAmbiente_Click);
            // 
            // groupApplicativi
            // 
            this.groupApplicativi.Items.Add(this.btnPrevisioneGas);
            this.groupApplicativi.Items.Add(this.btnUnitCommitment);
            this.groupApplicativi.Items.Add(this.btnPrezziMSD);
            this.groupApplicativi.Items.Add(this.btnValidazioneTL);
            this.groupApplicativi.Items.Add(this.btnPrevisioneCT);
            this.groupApplicativi.Items.Add(this.btnProgrammazioneImpianti);
            this.groupApplicativi.Items.Add(this.btnOfferteMGP);
            this.groupApplicativi.Items.Add(this.btnOfferteMSD);
            this.groupApplicativi.Items.Add(this.btnOfferteMB);
            this.groupApplicativi.Items.Add(this.btnInvioProgrammi);
            this.groupApplicativi.Items.Add(this.btnSistemaComandi);
            this.groupApplicativi.Label = "Applicativi";
            this.groupApplicativi.Name = "groupApplicativi";
            // 
            // btnPrevisioneGas
            // 
            this.btnPrevisioneGas.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPrevisioneGas.Image = ((System.Drawing.Image)(resources.GetObject("btnPrevisioneGas.Image")));
            this.btnPrevisioneGas.Label = "Previsione Gas";
            this.btnPrevisioneGas.Name = "btnPrevisioneGas";
            this.btnPrevisioneGas.ShowImage = true;
            this.btnPrevisioneGas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnUnitCommitment
            // 
            this.btnUnitCommitment.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUnitCommitment.Image = ((System.Drawing.Image)(resources.GetObject("btnUnitCommitment.Image")));
            this.btnUnitCommitment.Label = "Unit Commitment";
            this.btnUnitCommitment.Name = "btnUnitCommitment";
            this.btnUnitCommitment.ShowImage = true;
            this.btnUnitCommitment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnPrezziMSD
            // 
            this.btnPrezziMSD.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPrezziMSD.Image = ((System.Drawing.Image)(resources.GetObject("btnPrezziMSD.Image")));
            this.btnPrezziMSD.Label = "Prezzi MSD";
            this.btnPrezziMSD.Name = "btnPrezziMSD";
            this.btnPrezziMSD.ShowImage = true;
            this.btnPrezziMSD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnValidazioneTL
            // 
            this.btnValidazioneTL.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnValidazioneTL.Image = ((System.Drawing.Image)(resources.GetObject("btnValidazioneTL.Image")));
            this.btnValidazioneTL.Label = "Validazione TL";
            this.btnValidazioneTL.Name = "btnValidazioneTL";
            this.btnValidazioneTL.ShowImage = true;
            this.btnValidazioneTL.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnPrevisioneCT
            // 
            this.btnPrevisioneCT.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPrevisioneCT.Image = ((System.Drawing.Image)(resources.GetObject("btnPrevisioneCT.Image")));
            this.btnPrevisioneCT.Label = "Previsione CT";
            this.btnPrevisioneCT.Name = "btnPrevisioneCT";
            this.btnPrevisioneCT.ShowImage = true;
            this.btnPrevisioneCT.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnProgrammazioneImpianti
            // 
            this.btnProgrammazioneImpianti.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnProgrammazioneImpianti.Image = ((System.Drawing.Image)(resources.GetObject("btnProgrammazioneImpianti.Image")));
            this.btnProgrammazioneImpianti.Label = "Progr. Impianti";
            this.btnProgrammazioneImpianti.Name = "btnProgrammazioneImpianti";
            this.btnProgrammazioneImpianti.ShowImage = true;
            this.btnProgrammazioneImpianti.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnOfferteMGP
            // 
            this.btnOfferteMGP.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOfferteMGP.Image = ((System.Drawing.Image)(resources.GetObject("btnOfferteMGP.Image")));
            this.btnOfferteMGP.Label = "Offerte MGP";
            this.btnOfferteMGP.Name = "btnOfferteMGP";
            this.btnOfferteMGP.ShowImage = true;
            this.btnOfferteMGP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnOfferteMSD
            // 
            this.btnOfferteMSD.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOfferteMSD.Image = ((System.Drawing.Image)(resources.GetObject("btnOfferteMSD.Image")));
            this.btnOfferteMSD.Label = "Offerte MSD";
            this.btnOfferteMSD.Name = "btnOfferteMSD";
            this.btnOfferteMSD.ShowImage = true;
            this.btnOfferteMSD.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnOfferteMB
            // 
            this.btnOfferteMB.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnOfferteMB.Image = ((System.Drawing.Image)(resources.GetObject("btnOfferteMB.Image")));
            this.btnOfferteMB.Label = "Offerte MB";
            this.btnOfferteMB.Name = "btnOfferteMB";
            this.btnOfferteMB.ShowImage = true;
            this.btnOfferteMB.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnInvioProgrammi
            // 
            this.btnInvioProgrammi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInvioProgrammi.Image = ((System.Drawing.Image)(resources.GetObject("btnInvioProgrammi.Image")));
            this.btnInvioProgrammi.Label = "Invio Programmi";
            this.btnInvioProgrammi.Name = "btnInvioProgrammi";
            this.btnInvioProgrammi.ShowImage = true;
            this.btnInvioProgrammi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // btnSistemaComandi
            // 
            this.btnSistemaComandi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSistemaComandi.Image = ((System.Drawing.Image)(resources.GetObject("btnSistemaComandi.Image")));
            this.btnSistemaComandi.Label = "Sistema Comandi";
            this.btnSistemaComandi.Name = "btnSistemaComandi";
            this.btnSistemaComandi.ShowImage = true;
            this.btnSistemaComandi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnProgrammi_Click);
            // 
            // groupErrori
            // 
            this.groupErrori.Items.Add(this.btnMostraErrorPane);
            this.groupErrori.Label = "Errori";
            this.groupErrori.Name = "groupErrori";
            // 
            // btnMostraErrorPane
            // 
            this.btnMostraErrorPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMostraErrorPane.Image = ((System.Drawing.Image)(resources.GetObject("btnMostraErrorPane.Image")));
            this.btnMostraErrorPane.Label = "Mostra pannello";
            this.btnMostraErrorPane.Name = "btnMostraErrorPane";
            this.btnMostraErrorPane.ShowImage = true;
            this.btnMostraErrorPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMostraErrorPane_Click);
            // 
            // groupMSD
            // 
            this.groupMSD.Items.Add(this.labelMSD);
            this.groupMSD.Items.Add(this.cmbMSD);
            this.groupMSD.Label = "MSD";
            this.groupMSD.Name = "groupMSD";
            // 
            // labelMSD
            // 
            this.labelMSD.Label = "Mercato MSD";
            this.labelMSD.Name = "labelMSD";
            // 
            // cmbMSD
            // 
            this.cmbMSD.Label = "Mercato MSD";
            this.cmbMSD.Name = "cmbMSD";
            this.cmbMSD.ShowLabel = false;
            this.cmbMSD.Text = null;
            // 
            // groupStagione
            // 
            this.groupStagione.Items.Add(this.labelStagione);
            this.groupStagione.Items.Add(this.cmbStagione);
            this.groupStagione.Label = "Parametri";
            this.groupStagione.Name = "groupStagione";
            // 
            // labelStagione
            // 
            this.labelStagione.Label = "Stagione";
            this.labelStagione.Name = "labelStagione";
            // 
            // cmbStagione
            // 
            ribbonDropDownItemImpl1.Label = "Primavera";
            ribbonDropDownItemImpl2.Label = "Estate";
            ribbonDropDownItemImpl3.Label = "Autunno";
            ribbonDropDownItemImpl4.Label = "Inverno";
            this.cmbStagione.Items.Add(ribbonDropDownItemImpl1);
            this.cmbStagione.Items.Add(ribbonDropDownItemImpl2);
            this.cmbStagione.Items.Add(ribbonDropDownItemImpl3);
            this.cmbStagione.Items.Add(ribbonDropDownItemImpl4);
            this.cmbStagione.Label = "Mercato MSD";
            this.cmbStagione.Name = "cmbStagione";
            this.cmbStagione.ShowLabel = false;
            this.cmbStagione.Text = null;
            // 
            // TabHome
            // 
            this.TabHome.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabHome.ControlId.OfficeId = "TabHome";
            this.TabHome.Label = "TabHome";
            this.TabHome.Name = "TabHome";
            // 
            // TabInsert
            // 
            this.TabInsert.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabInsert.ControlId.OfficeId = "TabInsert";
            this.TabInsert.Label = "TabInsert";
            this.TabInsert.Name = "TabInsert";
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
            // TabSmartArtToolsDesign
            // 
            this.TabSmartArtToolsDesign.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabSmartArtToolsDesign.ControlId.OfficeId = "TabSmartArtToolsDesign";
            this.TabSmartArtToolsDesign.Label = "TabSmartArtToolsDesign";
            this.TabSmartArtToolsDesign.Name = "TabSmartArtToolsDesign";
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
            this.Tabs.Add(this.TabSmartArtToolsDesign);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ToolsExcelRibbon_Load);
            this.FrontOffice.ResumeLayout(false);
            this.FrontOffice.PerformLayout();
            this.groupChiudi.ResumeLayout(false);
            this.groupChiudi.PerformLayout();
            this.groupEmergenza.ResumeLayout(false);
            this.groupEmergenza.PerformLayout();
            this.groupConfigura.ResumeLayout(false);
            this.groupConfigura.PerformLayout();
            this.groupCalendario.ResumeLayout(false);
            this.groupCalendario.PerformLayout();
            this.groupModifica.ResumeLayout(false);
            this.groupModifica.PerformLayout();
            this.groupAggiorna.ResumeLayout(false);
            this.groupAggiorna.PerformLayout();
            this.groupAzioni.ResumeLayout(false);
            this.groupAzioni.PerformLayout();
            this.groupAmbienti.ResumeLayout(false);
            this.groupAmbienti.PerformLayout();
            this.groupApplicativi.ResumeLayout(false);
            this.groupApplicativi.PerformLayout();
            this.groupErrori.ResumeLayout(false);
            this.groupErrori.PerformLayout();
            this.groupMSD.ResumeLayout(false);
            this.groupMSD.PerformLayout();
            this.groupStagione.ResumeLayout(false);
            this.groupStagione.PerformLayout();
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
            this.TabSmartArtToolsDesign.ResumeLayout(false);
            this.TabSmartArtToolsDesign.PerformLayout();

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
        public Microsoft.Office.Tools.Ribbon.RibbonTab FrontOffice;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAggiorna;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAggiornaStruttura;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCalendar;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRampe;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupChiudi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChiudi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAggiornaDati;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAzioni;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAzioni;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupModifica;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnModifica;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAmbienti;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnProduzione;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnTest;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnDev;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOttimizza;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConfigura;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCalendario;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabSmartArtToolsDesign;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupConfigura;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnPrevisioneGas;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnUnitCommitment;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnPrezziMSD;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupApplicativi;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnValidazioneTL;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnPrevisioneCT;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnProgrammazioneImpianti;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnOfferteMGP;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnOfferteMSD;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnOfferteMB;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnInvioProgrammi;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnSistemaComandi;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnForzaEmergenza;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupErrori;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnMostraErrorPane;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConfiguraParametri;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupMSD;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelMSD;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbMSD;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupStagione;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel labelStagione;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cmbStagione;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupEmergenza;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEsportaXML;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnImportaXML;
    }

    partial class ThisRibbonCollection
    {
        internal ToolsExcelRibbon ToolsExcelRibbon
        {
            get { return this.GetRibbon<ToolsExcelRibbon>(); }
        }
    }
}
