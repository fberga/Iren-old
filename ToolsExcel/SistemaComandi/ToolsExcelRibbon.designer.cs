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
            this.FrontOffice = this.Factory.CreateRibbonTab();
            this.groupChiudi = this.Factory.CreateRibbonGroup();
            this.btnChiudi = this.Factory.CreateRibbonButton();
            this.btnForzaEmergenza = this.Factory.CreateRibbonToggleButton();
            this.groupConfigura = this.Factory.CreateRibbonGroup();
            this.btnConfigura = this.Factory.CreateRibbonButton();
            this.groupCalendario = this.Factory.CreateRibbonGroup();
            this.btnCalendar = this.Factory.CreateRibbonButton();
            this.groupAggiorna = this.Factory.CreateRibbonGroup();
            this.btnAggiornaDati = this.Factory.CreateRibbonButton();
            this.btnAggiornaStruttura = this.Factory.CreateRibbonButton();
            this.groupAzioni = this.Factory.CreateRibbonGroup();
            this.btnAzioni = this.Factory.CreateRibbonButton();
            this.btnOttimizza = this.Factory.CreateRibbonButton();
            this.btnRampe = this.Factory.CreateRibbonButton();
            this.groupModifica = this.Factory.CreateRibbonGroup();
            this.btnModifica = this.Factory.CreateRibbonToggleButton();
            this.groupAmbienti = this.Factory.CreateRibbonGroup();
            this.btnProduzione = this.Factory.CreateRibbonToggleButton();
            this.btnTest = this.Factory.CreateRibbonToggleButton();
            this.btnDev = this.Factory.CreateRibbonToggleButton();
            this.groupFileRete = this.Factory.CreateRibbonGroup();
            this.btnPrevisioneGas = this.Factory.CreateRibbonToggleButton();
            this.btnUnitCommitment = this.Factory.CreateRibbonToggleButton();
            this.btnPrezziMSD = this.Factory.CreateRibbonToggleButton();
            this.groupFileLocali = this.Factory.CreateRibbonGroup();
            this.btnValidazioneTL = this.Factory.CreateRibbonToggleButton();
            this.btnPrevisioneCT = this.Factory.CreateRibbonToggleButton();
            this.btnProgrammazioneImpianti = this.Factory.CreateRibbonToggleButton();
            this.btnOfferteMGP = this.Factory.CreateRibbonToggleButton();
            this.btnOfferteMSD = this.Factory.CreateRibbonToggleButton();
            this.btnOfferteMB = this.Factory.CreateRibbonToggleButton();
            this.btnInvioProgrammi = this.Factory.CreateRibbonToggleButton();
            this.btnSistemaComandi = this.Factory.CreateRibbonToggleButton();
            this.TabAddIns = this.Factory.CreateRibbonTab();
            this.TabHome = this.Factory.CreateRibbonTab();
            this.TabInsert = this.Factory.CreateRibbonTab();
            this.TabPageLayoutExcel = this.Factory.CreateRibbonTab();
            this.TabFormulas = this.Factory.CreateRibbonTab();
            this.TabData = this.Factory.CreateRibbonTab();
            this.TabReview = this.Factory.CreateRibbonTab();
            this.TabView = this.Factory.CreateRibbonTab();
            this.TabDeveloper = this.Factory.CreateRibbonTab();
            this.TabPrintPreview = this.Factory.CreateRibbonTab();
            this.TabBackgroundRemoval = this.Factory.CreateRibbonTab();
            this.TabSmartArtToolsDesign = this.Factory.CreateRibbonTab();
            this.groupErrori = this.Factory.CreateRibbonGroup();
            this.btnMostraErrorPane = this.Factory.CreateRibbonButton();
            this.FrontOffice.SuspendLayout();
            this.groupChiudi.SuspendLayout();
            this.groupConfigura.SuspendLayout();
            this.groupCalendario.SuspendLayout();
            this.groupAggiorna.SuspendLayout();
            this.groupAzioni.SuspendLayout();
            this.groupModifica.SuspendLayout();
            this.groupAmbienti.SuspendLayout();
            this.groupFileRete.SuspendLayout();
            this.groupFileLocali.SuspendLayout();
            this.TabAddIns.SuspendLayout();
            this.TabHome.SuspendLayout();
            this.TabInsert.SuspendLayout();
            this.TabPageLayoutExcel.SuspendLayout();
            this.TabFormulas.SuspendLayout();
            this.TabData.SuspendLayout();
            this.TabReview.SuspendLayout();
            this.TabView.SuspendLayout();
            this.TabDeveloper.SuspendLayout();
            this.TabPrintPreview.SuspendLayout();
            this.TabBackgroundRemoval.SuspendLayout();
            this.TabSmartArtToolsDesign.SuspendLayout();
            this.groupErrori.SuspendLayout();
            // 
            // FrontOffice
            // 
            this.FrontOffice.Groups.Add(this.groupChiudi);
            this.FrontOffice.Groups.Add(this.groupConfigura);
            this.FrontOffice.Groups.Add(this.groupCalendario);
            this.FrontOffice.Groups.Add(this.groupAggiorna);
            this.FrontOffice.Groups.Add(this.groupAzioni);
            this.FrontOffice.Groups.Add(this.groupModifica);
            this.FrontOffice.Groups.Add(this.groupAmbienti);
            this.FrontOffice.Groups.Add(this.groupFileRete);
            this.FrontOffice.Groups.Add(this.groupFileLocali);
            this.FrontOffice.Groups.Add(this.groupErrori);
            this.FrontOffice.Label = "Front Office";
            this.FrontOffice.Name = "FrontOffice";
            // 
            // groupChiudi
            // 
            this.groupChiudi.Items.Add(this.btnChiudi);
            this.groupChiudi.Items.Add(this.btnForzaEmergenza);
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
            // btnForzaEmergenza
            // 
            this.btnForzaEmergenza.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnForzaEmergenza.Image = ((System.Drawing.Image)(resources.GetObject("btnForzaEmergenza.Image")));
            this.btnForzaEmergenza.Label = "Forza Emergenza";
            this.btnForzaEmergenza.Name = "btnForzaEmergenza";
            this.btnForzaEmergenza.ShowImage = true;
            this.btnForzaEmergenza.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnForzaEmergenza_Click);
            // 
            // groupConfigura
            // 
            this.groupConfigura.Items.Add(this.btnConfigura);
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
            this.btnOttimizza.Label = "Esegui Ottimizzazione";
            this.btnOttimizza.Name = "btnOttimizza";
            this.btnOttimizza.ShowImage = true;
            this.btnOttimizza.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOttimizza_Click);
            // 
            // btnRampe
            // 
            this.btnRampe.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRampe.Image = ((System.Drawing.Image)(resources.GetObject("btnRampe.Image")));
            this.btnRampe.Label = "Seleziona Rampa";
            this.btnRampe.Name = "btnRampe";
            this.btnRampe.ShowImage = true;
            this.btnRampe.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRampe_Click);
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
            // groupAmbienti
            // 
            this.groupAmbienti.Items.Add(this.btnProduzione);
            this.groupAmbienti.Items.Add(this.btnTest);
            this.groupAmbienti.Items.Add(this.btnDev);
            this.groupAmbienti.Label = "Ambienti";
            this.groupAmbienti.Name = "groupAmbienti";
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
            // groupFileRete
            // 
            this.groupFileRete.Items.Add(this.btnPrevisioneGas);
            this.groupFileRete.Items.Add(this.btnUnitCommitment);
            this.groupFileRete.Items.Add(this.btnPrezziMSD);
            this.groupFileRete.Label = "File in rete";
            this.groupFileRete.Name = "groupFileRete";
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
            // groupFileLocali
            // 
            this.groupFileLocali.Items.Add(this.btnValidazioneTL);
            this.groupFileLocali.Items.Add(this.btnPrevisioneCT);
            this.groupFileLocali.Items.Add(this.btnProgrammazioneImpianti);
            this.groupFileLocali.Items.Add(this.btnOfferteMGP);
            this.groupFileLocali.Items.Add(this.btnOfferteMSD);
            this.groupFileLocali.Items.Add(this.btnOfferteMB);
            this.groupFileLocali.Items.Add(this.btnInvioProgrammi);
            this.groupFileLocali.Items.Add(this.btnSistemaComandi);
            this.groupFileLocali.Label = "File in locale";
            this.groupFileLocali.Name = "groupFileLocali";
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
            // TabAddIns
            // 
            this.TabAddIns.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabAddIns.Label = "TabAddIns";
            this.TabAddIns.Name = "TabAddIns";
            this.TabAddIns.Visible = false;
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
            // groupErrori
            // 
            this.groupErrori.Items.Add(this.btnMostraErrorPane);
            this.groupErrori.Label = "Errori";
            this.groupErrori.Name = "groupErrori";
            // 
            // btnMostraErrorPane
            // 
            this.btnMostraErrorPane.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnMostraErrorPane.Label = "Mostra pannello";
            this.btnMostraErrorPane.Name = "btnMostraErrorPane";
            this.btnMostraErrorPane.ShowImage = true;
            this.btnMostraErrorPane.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnMostraErrorPane_Click);
            // 
            // ToolsExcelRibbon
            // 
            this.Name = "ToolsExcelRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
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
            this.groupConfigura.ResumeLayout(false);
            this.groupConfigura.PerformLayout();
            this.groupCalendario.ResumeLayout(false);
            this.groupCalendario.PerformLayout();
            this.groupAggiorna.ResumeLayout(false);
            this.groupAggiorna.PerformLayout();
            this.groupAzioni.ResumeLayout(false);
            this.groupAzioni.PerformLayout();
            this.groupModifica.ResumeLayout(false);
            this.groupModifica.PerformLayout();
            this.groupAmbienti.ResumeLayout(false);
            this.groupAmbienti.PerformLayout();
            this.groupFileRete.ResumeLayout(false);
            this.groupFileRete.PerformLayout();
            this.groupFileLocali.ResumeLayout(false);
            this.groupFileLocali.PerformLayout();
            this.TabAddIns.ResumeLayout(false);
            this.TabAddIns.PerformLayout();
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
            this.TabPrintPreview.ResumeLayout(false);
            this.TabPrintPreview.PerformLayout();
            this.TabBackgroundRemoval.ResumeLayout(false);
            this.TabBackgroundRemoval.PerformLayout();
            this.TabSmartArtToolsDesign.ResumeLayout(false);
            this.TabSmartArtToolsDesign.PerformLayout();
            this.groupErrori.ResumeLayout(false);
            this.groupErrori.PerformLayout();

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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFileRete;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnPrevisioneGas;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnUnitCommitment;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnPrezziMSD;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFileLocali;
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
    }

    partial class ThisRibbonCollection
    {
        internal ToolsExcelRibbon ToolsExcelRibbon
        {
            get { return this.GetRibbon<ToolsExcelRibbon>(); }
        }
    }
}
