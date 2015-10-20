using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System.Data;
using System.Globalization;
using System.Reflection;
using Iren.ToolsExcel.Forms;
using System.Collections;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Tools;
using System.Configuration;

namespace ProvaRibbon
{
    public partial class AutoRibbon
    {
        #region Variabili

        /// <summary>
        /// Lista dei controlli (solo button e togglebutton).
        /// </summary>
        private ControlCollection _controls;
        /// <summary>
        /// lista degli id dei tasti abilitati.
        /// </summary>
        private List<string> _enabledControls = new List<string>();
        /// <summary>
        /// Indica se tutti i tasti (a parte Aggiorna Struttura) sono disabilitati.
        /// </summary>
        private bool _allDisabled = false;
        /// <summary>
        /// Componente da aggiungere all'actionsPane del documento.
        /// </summary>
        private ErrorPane _errorPane = new ErrorPane();
        /// <summary>
        /// Variabile per svolgere delle azioni custom coi ceck.
        /// </summary>
        private Check _checkFunctions = new Check();
        /// <summary>
        /// Classe per l'aggiunta di azioni custom dopo la modifica di un Range.
        /// </summary>
        public Modifica _modificaCustom = new Modifica();

        #endregion

        #region Proprietà

        /// <summary>
        /// Proprietà che permette l'indicizzazione per nome dei vari tasti della barra Ribbon. 
        /// La necessità di questa proprietà deriva dalla necessità di abilitare/disabilitare/nascondere i tasti leggendo i parametri del DB.
        /// </summary>
        public ControlCollection Controls { get; private set; }
        public List<RibbonGroup> Groups { get; private set; }

        #endregion

        public void InitializeComponent2()
        {
            //EventInfo ei = btnCalendar.GetType().GetEvent("Click");
            //MethodInfo hi = GetType().GetMethod("btnCalendar_Click", BindingFlags.Instance | BindingFlags.NonPublic);
            //Delegate d = Delegate.CreateDelegate(ei.EventHandlerType, null, hi);
            //ei.AddEventHandler(btnCalendar, d);


            //this.groupChiudi = this.Factory.CreateRibbonGroup();
            //this.btnEsportaXML = this.Factory.CreateRibbonButton();
            //this.btnValidazioneTL = this.Factory.CreateRibbonToggleButton();
            //this.labelMSD = this.Factory.CreateRibbonLabel();
            //this.cmbMSD = this.Factory.CreateRibbonComboBox();

            //this.btnChiudi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            //this.btnChiudi.Image = ToolsExcel.Base.Properties.Resources.chiudi;
            //this.btnChiudi.Label = "Chiudi";
            //this.btnChiudi.Name = "btnChiudi";
            //this.btnChiudi.ShowImage = true;

            DataBase.InitNewDB("Dev");

            if (DataBase.OpenConnection())
            {
                //TODO salvare anche questo negli XML
                DataTable dt = DataBase.Select("RIBBON.spGruppoControllo", "@IdApplicazione=" + ConfigurationManager.AppSettings["AppID"] + ";@IdUtente=62");

                Microsoft.Office.Tools.Ribbon.RibbonGroup grp = this.Factory.CreateRibbonGroup();
                Groups = new List<RibbonGroup>();
                
                int idGruppo = -1;

                foreach (DataRow r in dt.Rows)
                {
                    if (!r["IdGruppo"].Equals(idGruppo))
                    {
                        idGruppo = (int)r["IdGruppo"];
                        grp = this.Factory.CreateRibbonGroup();
                        grp.Name = r["NomeGruppo"].ToString();
                        grp.Label = r["LabelGruppo"].ToString();

                        this.FrontOffice.Groups.Add(grp);
                        Groups.Add(grp);
                    }

                    RibbonControl ctrl = null;

                    if(typeof(RibbonButton).FullName.Equals(r["SiglaTipologiaControllo"])) 
                    {
                        RibbonButton newBtn = this.Factory.CreateRibbonButton();

                        newBtn.ControlSize = (Microsoft.Office.Core.RibbonControlSize)r["ControlSize"];
                        newBtn.Image = (System.Drawing.Image)Iren.ToolsExcel.Base.Properties.Resources.ResourceManager.GetObject(r["Immagine"].ToString());
                        newBtn.Label = r["Label"].ToString();
                        newBtn.Name = r["Nome"].ToString();
                        newBtn.Description = r["Descrizione"].ToString();
                        newBtn.ScreenTip = r["ScreenTip"].ToString();
                        newBtn.ShowImage = true;
                        grp.Items.Add(newBtn);
                        ctrl = newBtn;
                    }
                    else if (typeof(RibbonToggleButton).FullName.Equals(r["SiglaTipologiaControllo"])) 
                    {
                        RibbonToggleButton newTglBtn = this.Factory.CreateRibbonToggleButton();

                        newTglBtn.ControlSize = (Microsoft.Office.Core.RibbonControlSize)r["ControlSize"];
                        newTglBtn.Image = (System.Drawing.Image)Iren.ToolsExcel.Base.Properties.Resources.ResourceManager.GetObject(r["Immagine"].ToString());
                        newTglBtn.Label = r["Label"].ToString();
                        newTglBtn.Name = r["Nome"].ToString();
                        newTglBtn.Description = r["Descrizione"].ToString();
                        newTglBtn.ScreenTip = r["ScreenTip"].ToString();
                        newTglBtn.ShowImage = true;

                        grp.Items.Add(newTglBtn);
                        ctrl = newTglBtn;
                    }
                    else if (typeof(RibbonComboBox).FullName.Equals(r["SiglaTipologiaControllo"])) 
                    {
                        RibbonLabel lb = this.Factory.CreateRibbonLabel();
                        lb.Label = r["Label"].ToString();
                        RibbonComboBox cmb = this.Factory.CreateRibbonComboBox();
                        cmb.ShowLabel = false;
                        cmb.Text = null;
                        cmb.Name = r["Nome"].ToString();

                        grp.Items.Add(lb);
                        grp.Items.Add(cmb);
                        ctrl = cmb;
                    }
                    ctrl.Enabled = r["Abilitato"].Equals("1");
                    //aggiungo l'evento al controllo appena creato
                    DataTable funzioni = DataBase.Select("RIBBON.spControlloFunzione", "@IdGruppoControllo=" + r["IdGruppoControllo"]);
                    foreach (DataRow f in funzioni.Rows)
                    {
                        EventInfo ei = ctrl.GetType().GetEvent(f["Evento"].ToString());
                        MethodInfo hi = GetType().GetMethod(f["NomeFunzione"].ToString(), BindingFlags.Instance | BindingFlags.NonPublic);
                        Delegate d = Delegate.CreateDelegate(ei.EventHandlerType, null, hi);
                        ei.AddEventHandler(ctrl, d);
                    }                    
                }

                Controls = new ControlCollection(this);
            }
        }

        #region Eventi

        /// <summary>
        /// Al caricamento del Ribbon imposta i tasti e la tab da visualizzare
        /// </summary>       
        private void AutoRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            //Initialize();
            //Workbook.ScreenUpdating = false;
            //Sheet.Protected = false;

            //forzo aggiornamento label iniziale
            //Iren.ToolsExcel.Utility.Workbook.AggiornaLabelStatoDB();

            //se non sono in debug toglie le intestazioni
#if !DEBUG
            foreach(Excel.Worksheet ws in Globals.ThisWorkbook.Sheets)
            {
                ws.Activate();
                Globals.ThisWorkbook.ThisApplication.ActiveWindow.DisplayHeadings = false;
            }
            Globals.Main.Activate();
#endif
            //se sono al primo avvio dopo il rilascio di un aggiornamento o il cambio di giorno/mercato aggiorno la struttura
            bool isUpdated = true;
            //if (Workbook.CategorySheets.Count == 0 || Repository.DaAggiornare)
            //{
            //    Aggiorna aggiorna = new Aggiorna();
            //    isUpdated = aggiorna.Struttura(avoidRepositoryUpdate: false);
            //}

            if (isUpdated)
            {
                ((RibbonButton)Controls["btnCalendario"]).Label = DataBase.DataAttiva.ToString("dddd dd MMM yyyy");

                //seleziono l'ambiente attivo
                ((RibbonToggleButton)Controls["btn" + DataBase.DB.Ambiente]).Checked = true;

                //RefreshChecks();

                //se esce con qualche errore il tasto mantiene lo stato a cui era impostato
                ((RibbonToggleButton)Controls["btnModifica"]).Checked = false;
                ((RibbonToggleButton)Controls["btnModifica"]).Image = Iren.ToolsExcel.Base.Properties.Resources.modificaNO;
                ((RibbonToggleButton)Controls["btnModifica"]).Label = "Modifica NO";
                try
                {
                    Sheet.AbilitaModifica(false);
                }
                catch { }

                //seleziono il tasto dell'applicativo aperto
                CheckTastoApplicativo();

                //aggiungo errorPane
                Globals.ThisWorkbook.ActionsPane.Controls.Add(_errorPane);
                Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane = false;
                Globals.ThisWorkbook.ActionsPane.AutoScroll = false;
                Globals.ThisWorkbook.ActionsPane.SizeChanged += ActionsPane_SizeChanged;

                //aggiungo un altro handler per cell click
                Globals.ThisWorkbook.SheetSelectionChange += CheckSelection;
                Globals.ThisWorkbook.SheetSelectionChange += Handler.SelectionClick;

                //aggiungo un handler per modificare lo stato dei tasti di export a seconda dello stato del DB
                DataBase.DB.PropertyChanged += StatoDB_Changed;
                StatoDB_Changed(null, null);
            }

            //Sheet.Protected = true;
            //Workbook.ScreenUpdating = true;
            //SplashScreen.Close();
        }

        private void StatoDB_Changed(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            if (DataBase.OpenConnection())
            {
                if (Controls["btnEsportaXML"].Enabled)
                    Controls["btnEsportaXML"].Enabled = false;
                if (Controls["btnImportaXML"].Enabled)
                    Controls["btnImportaXML"].Enabled = false;

                if (_enabledControls.Contains("btnProduzione"))
                    Controls["btnProduzione"].Enabled = true;
                if (_enabledControls.Contains("btnTest"))
                    Controls["btnTest"].Enabled = true;
                if (_enabledControls.Contains("btnDev"))
                    Controls["btnDev"].Enabled = true;
                if (_enabledControls.Contains("btnAggiornaDati"))
                    Controls["btnAggiornaDati"].Enabled = true;
                if (_enabledControls.Contains("btnAggiornaStruttura"))
                    Controls["btnAggiornaStruttura"].Enabled = true;
                if (_enabledControls.Contains("btnConfiguraParametri"))
                    Controls["btnConfiguraParametri"].Enabled = true;

                DataBase.CloseConnection();
            }
            else
            {
                if (_enabledControls.Contains("btnEsportaXML"))
                    Controls["btnEsportaXML"].Enabled = true;
                if (_enabledControls.Contains("btnImportaXML"))
                    Controls["btnImportaXML"].Enabled = true;

                if (Controls["btnProduzione"].Enabled)
                    Controls["btnProduzione"].Enabled = false;
                if (Controls["btnTest"].Enabled)
                    Controls["btnTest"].Enabled = false;
                if (Controls["btnDev"].Enabled)
                    Controls["btnDev"].Enabled = false;
                if (Controls["btnAggiornaDati"].Enabled)
                    Controls["btnAggiornaDati"].Enabled = false;
                if (Controls["btnAggiornaStruttura"].Enabled)
                    Controls["btnAggiornaStruttura"].Enabled = false;
                if (Controls["btnConfiguraParametri"].Enabled)
                    Controls["btnConfiguraParametri"].Enabled = false;
            }
        }
        /// <summary>
        /// Handler del click sul tasto di configurazione dei parametri. Apre il form che permette di modificare i valori dei parametri definiti per il foglio.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfiguraParametri_Click(object sender, RibbonControlEventArgs e)
        {
            FormModificaParametri form = new FormModificaParametri();
            if (!form.IsDisposed)
                form.Show();
        }
        /// <summary>
        /// Handler del SheetSelectionChange. Funzione che controlla se la cella selezionata è un Check. Si trova qui e non dentro la Classe Base.Handler perché deve interagire con l'errorPane 
        /// (non è possibile farlo dal namespace Base in quanto si creerebbe uno using circolare)
        /// </summary>
        /// <param name="Sh">Worksheet dove è stato eseguito il cambio di selezione</param>
        /// <param name="Target">Range dove è stato eseguito il cambio di selezione</param>
        private void CheckSelection(object Sh, Excel.Range Target)
        {
            try
            {
                if (!Workbook.FromErrorPane)
                {
                    DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.Check);
                    Range rng = new Range(Target.Row, Target.Column);
                    if (definedNames.HasCheck() && definedNames.IsCheck(rng))
                    {
                        Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane = true;
                        _errorPane.SelectNode("'" + Target.Worksheet.Name + "'!" + rng.ToString());

                    }
                }
            }
            catch { }
        }
        /// <summary>
        /// Handler del ridimensionamento dell'ActionsPane del foglio, ridimensiona il componente ErrorPane in modo da adattarlo alle nuove dimensioni.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ActionsPane_SizeChanged(object sender, EventArgs e)
        {
            _errorPane.SetDimension(Globals.ThisWorkbook.ActionsPane.Width, Globals.ThisWorkbook.ActionsPane.Height);
        }
        /// <summary>
        /// Handler del click sui toggle buttons di cambio ambiente selezionato. Cambia la selezione, fa il refresh del file di configurazione e attiva l'aggiornamento della struttura del foglio.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSelezionaAmbiente_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton ambienteScelto = (RibbonToggleButton)sender;

            if (DataBase.OpenConnection())
            {
                int count = 0;
                foreach (RibbonToggleButton button in FrontOffice.Groups.First(g => g.Label.ToLower() == "ambienti").Items)
                {
                    if (button.Checked)
                    {
                        button.Checked = false;
                        count++;
                    }
                }
                //se maggiore di 1 allora c'è un cambio ambiente altrimenti doppio click sullo stesso e non faccio nulla
                if (count > 1)
                {
                    Workbook.InsertLog(Iren.ToolsExcel.Core.DataBase.TipologiaLOG.LogModifica, "Attivato ambiente " + ambienteScelto.Label);
                    DataBase.SwitchEnvironment(ambienteScelto.Label);

                    btnAggiornaStruttura_Click(null, null);
                }

                ambienteScelto.Checked = true;
                DataBase.CloseConnection();
            }
            else
            {
                ambienteScelto.Checked = false;

                System.Windows.Forms.MessageBox.Show("Non è possibile effettuare un cambio di ambiente quando il sistema è in emergenza...", Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);
            }
        }
        /// <summary>
        /// Handler del click del tasto di aggiornamento della struttura. Avvisa l'utente ed esegue l'aggiornamento della struttura. Esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAggiornaStruttura_Click(object sender, RibbonControlEventArgs e)
        {
            //avviso all'utente
            var response = System.Windows.Forms.DialogResult.Yes;

            if (sender != null && e != null)
                response = System.Windows.Forms.MessageBox.Show("Eseguire l'aggiornamento della struttura?", Simboli.nomeApplicazione + " - ATTENZIONE!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);

            if (response == System.Windows.Forms.DialogResult.Yes)
            {
                Workbook.ScreenUpdating = false;
                Sheet.Protected = false;

                Aggiorna aggiorna = new Aggiorna();
                if (aggiorna.Struttura(avoidRepositoryUpdate: false))
                    Workbook.InsertLog(Iren.ToolsExcel.Core.DataBase.TipologiaLOG.LogModifica, "Aggiorna struttura");

                RefreshChecks();

                Sheet.Protected = true;
                Workbook.ScreenUpdating = true;

                if (_allDisabled)
                    AbilitaTasti(true);
            }
        }
        /// <summary>
        /// Handler del click del tasto di cambio data. Verifica che la data selezionata sia diversa da quella attuale e fa partire il controllo per vedere se ci siano modifiche alla struttura attraverso
        /// DataBase.SP.CHECKMODIFICASTRUTTURA. Se ci sono aggiorno la struttra, altrimenti aggiorno semplicemente i dati. Esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            //apro il form calendario
            Iren.ToolsExcel.Forms.FormCalendar cal = new FormCalendar();

            cal.Top = System.Windows.Forms.Cursor.Position.Y - 20;
            cal.Left = System.Windows.Forms.Cursor.Position.X - 20;

            DateTime calDate = cal.ShowDialog();
            cal.Dispose();
            Workbook.Application.Windows[1].Activate();
            //verifico che la data sia stata cambiata
            if (calDate != DataBase.DataAttiva)
            {
                //Workbook.ScreenUpdating = false;
                Sheet.Protected = false;
                SplashScreen.Show();

                Workbook.ChangeAppSettings("DataAttiva", calDate.ToString("yyyyMMdd"));
                ((RibbonButton)sender).Label = calDate.ToString("dddd dd MMM yyyy");

                Aggiorna aggiorna = new Aggiorna();
                if (DataBase.OpenConnection())
                {
                    Workbook.InsertLog(Iren.ToolsExcel.Core.DataBase.TipologiaLOG.LogModifica, "Cambio Data a " + ((RibbonButton)sender).Label);
                    DataBase.ChangeDate(calDate);
                    DataBase.ExecuteSPApplicazioneInit();

                    DataTable stato = DataBase.Select(DataBase.SP.CHECKMODIFICASTRUTTURA, "@DataOld=" + DataBase.DataAttiva.ToString("yyyyMMdd") + ";@DataNew=" + calDate.ToString("yyyyMMdd"));

                    if (stato != null && stato.Rows.Count > 0 && stato.Rows[0]["Stato"].Equals(1))
                        aggiorna.Struttura(avoidRepositoryUpdate: false);
                    else
                        aggiorna.Dati();

                    Workbook.RefreshLog();
                }
                else  //emergenza
                {
                    DataBase.ChangeDate(calDate);
                    aggiorna.Emergenza();
                }

                RefreshChecks();

                Sheet.Protected = true;
                Workbook.ScreenUpdating = true;
            }
        }
        /// <summary>
        /// Handler del click del tasto di selezione rampe. Apre il form per la selezione delle rampe ed esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRampe_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            //prendo il nome sheet e il range selezionato (per poter lavorare su più giorni nel caso ci fosse necessità)
            string sheet = Workbook.ActiveSheet.Name;
            Excel.Range rng = Workbook.Application.Selection;

            DefinedNames definedNames = new DefinedNames(sheet);
            FormSelezioneUP selUP = new FormSelezioneUP("PQNR_PROFILO");

            //controllo se nel range selezionato è definita un'entità
            if (sheet == "Iren Termo" && definedNames.IsDefined(rng.Row))
            {
                string nome = definedNames.GetNameByAddress(rng.Row, rng.Column);
                string siglaEntita = nome.Split(Simboli.UNION[0])[0];

                //controllo se l'entità ha la possibilità di selezionare le rampe
                DataView entitaInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
                entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaInformazione = 'PQNR_PROFILO' AND IdApplicazione = " + Simboli.AppID;

                if (entitaInformazioni.Count == 0)
                {
                    //avviso l'utente che l'entità selezionata non ha l'opzione
                    if (System.Windows.Forms.MessageBox.Show("L'operazione selezionata non è disponibile per l'UP selezionata, selezionarne un'altra dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes
                        && selUP.ShowDialog().ToString() != "")
                    {
                        Iren.ToolsExcel.Forms.FormRampe rampe = new FormRampe(Workbook.Application.Selection);
                        rampe.ShowDialog();
                        rampe.Dispose();
                    }
                }
                else
                {
                    Iren.ToolsExcel.Forms.FormRampe rampe = new FormRampe(Workbook.Application.Selection);
                    rampe.ShowDialog();
                    rampe.Dispose();
                }
            }
            //sono in un foglio diverso da Iren Termo o su una cella senza definizione di nomi
            else if (System.Windows.Forms.MessageBox.Show("Nessuna UP selezionata, selezionarne una dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes
                && selUP.ShowDialog().ToString() != "")
            {
                Iren.ToolsExcel.Forms.FormRampe rampe = new FormRampe(Workbook.Application.Selection);
                rampe.ShowDialog();
                rampe.Dispose();
            }
            selUP.Dispose();
            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di aggiornamento dei dati. Aziona la funzione AggiornaDati ed esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAggiornaDati_Click(object sender, RibbonControlEventArgs e)
        {
            var response = System.Windows.Forms.MessageBox.Show("Eseguire l'aggiornamento dei dati?", Simboli.nomeApplicazione + " - ATTENZIONE!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
            if (response == System.Windows.Forms.DialogResult.Yes)
            {
                Workbook.ScreenUpdating = false;
                Sheet.Protected = false;

                Aggiorna aggiorna = new Aggiorna();
                if (aggiorna.Dati())
                    Workbook.InsertLog(Iren.ToolsExcel.Core.DataBase.TipologiaLOG.LogModifica, "Aggiorna Dati");

                RefreshChecks();

                Sheet.Protected = true;
                Workbook.ScreenUpdating = true;
            }
        }
        /// <summary>
        /// Handler del click del tasto delle azioni. Mostra il form delle azioni ed esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAzioni_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            FormAzioni frmAz = new FormAzioni(new Esporta(), new Riepilogo(), new Carica());
            frmAz.ShowDialog();

            RefreshChecks();

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di modifica. Attiva e disattiva la modifica foglio. Nel caso di disattivazione, aggiorna i check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnModifica_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.Application.EnableEvents = true;
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            Simboli.ModificaDati = ((RibbonToggleButton)sender).Checked;

            if (((RibbonToggleButton)sender).Checked)
            {
                AbilitaTasti(false);
                ((RibbonToggleButton)sender).Enabled = true;
                ((RibbonToggleButton)sender).Image = Iren.ToolsExcel.Base.Properties.Resources.modificaSI;
                ((RibbonToggleButton)sender).Label = "Modifica SI";
                Workbook.WB.SheetChange += Handler.StoreEdit;
                //Aggiungo handler per azioni custom nel caso servisse
                Workbook.WB.SheetChange += _modificaCustom.Range;
            }
            else
            {
                RefreshChecks();
                //salva modifiche sul db
                Sheet.SalvaModifiche();
                DataBase.SalvaModificheDB();
                ((RibbonToggleButton)sender).Image = Iren.ToolsExcel.Base.Properties.Resources.modificaNO;
                ((RibbonToggleButton)sender).Label = "Modifica NO";
                Workbook.WB.SheetChange -= Handler.StoreEdit;
                //Rimuovo handler per azioni custom nel caso servisse
                Workbook.WB.SheetChange -= _modificaCustom.Range;

                //aggiorno i label dello stato nel caso sia necessario!
                Workbook.AggiornaLabelStatoDB();

                AbilitaTasti(true);
                //disabilito i tasti legati alla connessione se necessario
                StatoDB_Changed(null, null);
            }
            Sheet.AbilitaModifica(((RibbonToggleButton)sender).Checked);

            Workbook.RefreshLog();

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di Ottimizzazione.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOttimizza_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            Excel.Range rng = Workbook.Application.Selection;

            DefinedNames definedNames = new DefinedNames(Workbook.ActiveSheet.Name, DefinedNames.InitType.Naming);

            //inizializzo ottimizzatore e il form di selezione entità per l'ottimo.
            Optimizer opt = new Optimizer();
            FormSelezioneUP selUP = new FormSelezioneUP("OTTIMO");

            if (definedNames.IsDefined(rng.Row))
            {
                string nome = definedNames.GetNameByAddress(rng.Row, rng.Column);
                string siglaEntita = nome.Split(Simboli.UNION[0])[0];

                DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Simboli.AppID;

                if (categoriaEntita.Count > 0)
                    siglaEntita = categoriaEntita[0]["Gerarchia"] is DBNull ? siglaEntita : categoriaEntita[0]["Gerarchia"].ToString();

                DataView entitaInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
                entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaInformazione = 'OTTIMO' AND IdApplicazione = " + Simboli.AppID;

                if (entitaInformazioni.Count == 0)
                {
                    if (System.Windows.Forms.MessageBox.Show("L'UP selezionata non può essere ottimizzata, selezionarne un'altra dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes)
                    {
                        siglaEntita = selUP.ShowDialog().ToString();
                        if (siglaEntita != null)
                        {
                            SplashScreen.Show();
                            SplashScreen.UpdateStatus("Ottimizzo " + siglaEntita);
                            opt.EseguiOttimizzazione(siglaEntita);
                            SplashScreen.UpdateStatus("Salvo modifiche su DB");
                            Sheet.SalvaModifiche();
                            DataBase.SalvaModificheDB();
                            SplashScreen.Close();
                        }
                    }
                }
                else
                {
                    SplashScreen.Show();
                    SplashScreen.UpdateStatus("Ottimizzo " + siglaEntita);
                    opt.EseguiOttimizzazione(siglaEntita);
                    SplashScreen.UpdateStatus("Salvo modifiche su DB");
                    Sheet.SalvaModifiche();
                    DataBase.SalvaModificheDB();
                    SplashScreen.Close();
                }
            }
            else if (System.Windows.Forms.MessageBox.Show("Nessuna UP selezionata, selezionarne una dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes)
            {
                object siglaEntita = selUP.ShowDialog();
                if (siglaEntita != null)
                {
                    SplashScreen.Show();
                    SplashScreen.UpdateStatus("Ottimizzo " + siglaEntita);
                    opt.EseguiOttimizzazione(siglaEntita);
                    SplashScreen.UpdateStatus("Salvo modifiche su DB");
                    Sheet.SalvaModifiche();
                    DataBase.SalvaModificheDB();
                    SplashScreen.Close();
                }
            }
            selUP.Dispose();
            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di modifica parametri. Mostra il form di modifica dei parametri utente.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfigura_Click(object sender, RibbonControlEventArgs e)
        {
            FormConfiguraPercorsi conf = new FormConfiguraPercorsi();
            conf.ShowDialog();
            conf.Dispose();
        }
        /// <summary>
        /// Handler del click dei tasti delle varie applicazioni. Abilita il foglio selezionato.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnProgrammi_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton btn = (RibbonToggleButton)sender;

            if (!btn.Checked)
            {
                btn.Checked = true;
            }
            else
            {
                btn.Checked = false;
                //TODO Controllare e cambiare il path
                Handler.SwitchWorksheet(btn.Name.Substring(3));
            }
        }
        /// <summary>
        /// Handler del click del tasto per forzare l'emergenza. Disabilita le connessioni al DB.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnForzaEmergenza_Click(object sender, RibbonControlEventArgs e)
        {
            Simboli.EmergenzaForzata = ((RibbonToggleButton)sender).Checked;
            StatoDB_Changed(null, null);
        }
        /// <summary>
        /// Handler del click del tasto di chiusura. Chiude l'applicativo.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChiudi_Click(object sender, RibbonControlEventArgs e)
        {
            //TextInfo ti = new CultureInfo("it-IT", false).TextInfo;
            //string pathStr = Iren.ToolsExcel.Utility.Workbook.GetUsrConfigElement("backup").Value;
            //if (!Directory.Exists(pathStr))
            //    Directory.CreateDirectory(pathStr);

            //string filename = ti.ToTitleCase(Simboli.nomeApplicazione).Replace(" ", "") + "_Backup_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsm";

            //Globals.ThisWorkbook.SaveCopyAs(Path.Combine(pathStr, filename));
            Globals.ThisWorkbook.ThisApplication.Quit();
        }
        /// <summary>
        /// Handler del click del tasto per visualizzare l'actionsPane del documento.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMostraErrorPane_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            if (!Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane)
                Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane = true;

            _errorPane.SetDimension(Globals.ThisWorkbook.ActionsPane.Width, Globals.ThisWorkbook.ActionsPane.Height);

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler della selezione di un nuovo mercato in cmbMSD su ribbon. Aggiorna la struttura dei fogli.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbMSD_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            //Simboli.AppID = Simboli.GetAppIDByMercato(cmbMSD.Text);
            Aggiorna aggiorna = new Aggiorna();
            aggiorna.Struttura(avoidRepositoryUpdate: true);

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler per il cambio di stagione da cmnStagione su ribbon. Imposta il valore della riga nascosta.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cmbStagione_TextChanged(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            Sheet.Protected = false;

            //Simboli.Stagione = cmbStagione.Text;

            Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
        }

        private void btnEsportaXML_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.ScreenUpdating = false;
            SplashScreen.Show();
            SplashScreen.UpdateStatus("Esporto tutte le informazioni del foglio");

            EsportaXML exp = new EsportaXML();
            exp.RunExport();
            SplashScreen.Close();
            Workbook.ScreenUpdating = true;
        }

        private void btnImportaXML_Click(object sender, RibbonControlEventArgs e)
        {
            FormImportXML frmXML = new FormImportXML();
            frmXML.ShowDialog();
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Funzione per aggiornate i check in seguito ad operazioni di modifica del foglio.
        /// </summary>
        private void RefreshChecks()
        {
            Workbook.ScreenUpdating = false;
            //Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            try
            {
                _errorPane.RefreshCheck(_checkFunctions);
            }
            catch { }
            //Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            //Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Metodo di inizializzazione della Tab Front Office. Visualizza e abilita i tasti in base alle specifiche da DB. Da notare che se ci sono aggiornamenti, bisogna caricare la struttura e riavviare l'applicativo.
        /// </summary>
        private void Initialize()
        {
            
            //DataView controlli = new DataView();

            //if (DataBase.OpenConnection())
            //{
            //    //Repository.CaricaApplicazioneRibbon();
            //    //controlli = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE_RIBBON].DefaultView;
            //    DataBase.CloseConnection();
            //}
            //else
            //{
            //    try
            //    {
            //        controlli = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE_RIBBON].DefaultView;
            //    }
            //    catch
            //    {
            //        controlli = new DataView();
            //    }
            //}

//            if (controlli.Count > 0)
//            {
//                foreach (DataRowView controllo in controlli)
//                {
//                    Controls[controllo["NomeControllo"].ToString()].Visible = controllo["Visibile"].Equals("1");
//                    Controls[controllo["NomeControllo"].ToString()].Enabled = controllo["Abilitato"].Equals("1");
//                    if (controllo["Abilitato"].Equals("1"))
//                        _enabledControls.Add(controllo["NomeControllo"].ToString());

//                    if (Controls[controllo["NomeControllo"].ToString()].GetType().ToString().Contains("ToggleButton"))
//                    {
//                        ((RibbonToggleButton)Controls[controllo["NomeControllo"].ToString()]).Checked = controllo["Stato"].Equals("1");
//                    }
//                }

//                List<RibbonGroup> groups = FrontOffice.Groups.ToList();
//                foreach (RibbonGroup group in groups)
//                    group.Visible = group.Items.Any(c => c.Visible);
//            }
//            else
//            {
//                foreach (RibbonControl control in Controls)
//                {
//#if !DEBUG
//                    control.Visible = true;
//                    control.Enabled = false;
//#else
//                    control.Visible = true;
//                    control.Enabled = true;
//#endif

//                    if (control.GetType().ToString().Contains("ToggleButton"))
//                        ((RibbonToggleButton)control).Checked = false;
//                }
//            }

            //ComboBox mercati
            //if (groupMSD.Visible)
            //{
            //    if (Workbook.AppSettings("Mercati") != null)
            //    {
            //        string[] mercati = Workbook.AppSettings("Mercati").Split('|');
            //        cmbMSD.Items.Clear();
            //        foreach (string mercato in mercati)
            //        {
            //            RibbonDropDownItem i = Factory.CreateRibbonDropDownItem();
            //            i.Label = mercato;
            //            cmbMSD.Items.Add(i);
            //        }

            //        cmbMSD.TextChanged -= cmbMSD_TextChanged;
            //        cmbMSD.Text = Simboli.Mercato;
            //        cmbMSD.TextChanged += cmbMSD_TextChanged;
            //    }
            //}

            //ComboBox stagioni
            //if (groupStagione.Visible)
            //{
            //    if (Workbook.AppSettings("Stagioni") != null)
            //    {
            //        string[] stagioni = Workbook.AppSettings("Stagioni").Split('|');
            //        cmbStagione.Items.Clear();
            //        foreach (string stagione in stagioni)
            //        {
            //            RibbonDropDownItem i = Factory.CreateRibbonDropDownItem();
            //            i.Label = stagione;
            //            cmbStagione.Items.Add(i);
            //        }

            //        cmbStagione.TextChanged -= cmbStagione_TextChanged;
            //        cmbStagione.Text = Simboli.Stagione;
            //        cmbStagione.TextChanged += cmbStagione_TextChanged;
            //    }
            //}
        }
        /// <summary>
        /// Metodo che seleziona il tasto corretto tra quelli degli applicativi presenti nella Tab Front Office. La selezione avviene in base all'ID applicazione scritto sul file di configurazione.
        /// </summary>
        private void CheckTastoApplicativo()
        {
            switch (ConfigurationManager.AppSettings["AppID"])
            {
                case "1":
                    ((RibbonToggleButton)Controls["btnOfferteMGP"]).Checked = true;
                    break;
                case "2":
                case "3":
                case "4":
                case "13":
                    ((RibbonToggleButton)Controls["btnInvioProgrammi"]).Checked = true;
                    break;
                case "5":
                    ((RibbonToggleButton)Controls["btnProgrammazioneImpianti"]).Checked = true;
                    break;
                case "6":
                    ((RibbonToggleButton)Controls["btnUnitCommitment"]).Checked = true;
                    break;
                case "7":
                    ((RibbonToggleButton)Controls["btnPrezziMSD"]).Checked = true;
                    break;
                case "8":
                    ((RibbonToggleButton)Controls["btnSistemaComandi"]).Checked = true;
                    break;
                case "9":
                    ((RibbonToggleButton)Controls["btnOfferteMSD"]).Checked = true;
                    break;
                case "10":
                    ((RibbonToggleButton)Controls["btnOfferteMB"]).Checked = true;
                    break;
                case "11":
                    ((RibbonToggleButton)Controls["btnValidazioneTL"]).Checked = true;
                    break;
                case "12":
                    ((RibbonToggleButton)Controls["btnPrevisioneCT"]).Checked = true;
                    break;
            }



        }
        /// <summary>
        /// Abilito tutti i tasti nel caso in cui, ad esempio in seguito a un rilascio, questi vengano disabilitati da DisabilitaTasti.
        /// </summary>
        private void AbilitaTasti(bool enable)
        {
            foreach (string control in _enabledControls)
                Controls[control].Enabled = enable;

            _allDisabled = enable;
        }

        #endregion
    }

    #region Controls Collection

    /// <summary>
    /// Classi che permettono di indicizzare per nome tutti i controlli contenuti nei gruppi della Tab Front Office
    /// </summary>
    public class ControlCollection : IEnumerable
    {
        #region Variabili

        private AutoRibbon _ribbon;
        private Dictionary<string, RibbonControl> _controls = new Dictionary<string, RibbonControl>();

        #endregion

        #region Proprietà

        public int Count
        {
            get { return _controls.Count; }
        }

        public RibbonControl this[int i]
        {
            get { return _controls.ElementAt(i).Value; }
        }

        public RibbonControl this[string name]
        {
            get { return _controls[name]; }
        }

        #endregion

        #region Metodi

        internal ControlCollection(AutoRibbon ribbon)
        {
            foreach (RibbonGroup group in ribbon.Groups)
                foreach (RibbonControl control in group.Items)
                    _controls.Add(control.Name, control);
        }

        public IEnumerator GetEnumerator()
        {
            return new ControlEnumerator(_ribbon);
        }

        public IEnumerable<KeyValuePair<string, RibbonControl>> AsEnumerable()
        {
            return _controls.AsEnumerable();
        }

        #endregion
    }
    public class ControlEnumerator : IEnumerator
    {
        #region Variabili

        private AutoRibbon _ribbon;
        private int _pos = -1;
        private int _max = -1;

        #endregion

        #region Costruttori

        public ControlEnumerator(AutoRibbon ribbon)
        {
            _ribbon = ribbon;
            _max = ribbon.Controls.Count;
        }

        #endregion

        #region Metodi

        public object Current
        {
            get { return _ribbon.Controls[_pos]; }
        }
        public bool MoveNext()
        {
            _pos++;
            return _pos < _max;
        }
        public void Reset()
        {
            _pos = -1;
        }

        #endregion
    }

    #endregion
}
