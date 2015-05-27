using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using Iren.ToolsExcel.Forms;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Collections.Generic;
using System.Collections;
using System.IO;

// ***************************************************** SISTEMA COMANDI ***************************************************** //

namespace Iren.ToolsExcel
{
    public partial class ToolsExcelRibbon
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

        #endregion

        #region Proprietà

        /// <summary>
        /// Proprietà che permette l'indicizzazione per nome dei vari tasti della barra Ribbon. 
        /// La necessità di questa proprietà deriva dalla necessità di abilitare/disabilitare/nascondere i tasti leggendo i parametri del DB.
        /// </summary>
        public ControlCollection Controls
        {
            get { return _controls; }
        }

        #endregion

        #region Eventi

        /// <summary>
        /// Al caricamento del Ribbon imposta i tasti e la tab da visualizzare
        /// </summary>       
        private void ToolsExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Initialize();

            //se non sono in debug toglie le intestazioni
#if !DEBUG
            foreach(Excel.Worksheet ws in Globals.ThisWorkbook.Sheets)
            {
                ws.Activate();
                Globals.ThisWorkbook.ThisApplication.ActiveWindow.DisplayHeadings = false;
            }
            Globals.Main.Activate();
#endif
            //se sono al primo avvio dopo un aggiornamento disabilito tutto a parte il tasto per aggiornare la struttura
            if (Workbook.WB.Sheets.Count <= 2)
                DisabilitaTasti();

            DateTime cfgDate = DateTime.ParseExact(Workbook.AppSettings("DataInizio"), "yyyyMMdd", CultureInfo.InvariantCulture);
            btnCalendar.Label = cfgDate.ToString("dddd dd MMM yyyy");

            //seleziono l'ambiente attivo
            ((RibbonToggleButton)Controls["btn" + Workbook.AppSettings("DB")]).Checked = true;

            //se esce con qualche errore il tasto mantiene lo stato a cui era impostato
            btnModifica.Checked = false;
            btnModifica.Image = Iren.ToolsExcel.Base.Properties.Resources.modificaNO_icon;
            btnModifica.Label = "Modifica NO";
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

            try
            {
                Sheet.Proteggi(false);
                _errorPane.RefreshCheck(_checkFunctions);
                Sheet.Proteggi(true);
            }
            catch { }

            //aggiungo un altro handler per cell click
            Globals.ThisWorkbook.SheetSelectionChange += CheckSelection;
            Globals.ThisWorkbook.SheetSelectionChange += SelectionClick;
        }

        private void SelectionClick(object Sh, Excel.Range Target)
        {
            DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.SelectionOnly);
            Range rng = new Range(Target.Row, Target.Column);
            SelectionObj sel;
            int val;
            if (definedNames.HasSelections() && definedNames.TryGetSelectionByPeer(rng, out sel, out val))
            {
                Target.Worksheet.Unprotect(Simboli.pwd);
                if (sel != null)
                {
                    Workbook.WB.SheetChange -= Handler.StoreEdit;
                    //Workbook.WB.Application.EnableEvents = false;
                    sel.ClearSelections(Target.Worksheet);
                    Target.Worksheet.Range[sel.GetByValue(val)].Value = "x";
                    //Workbook.WB.Application.EnableEvents = true;
                    Workbook.WB.SheetChange += Handler.StoreEdit;
                    Target.Worksheet.Range[sel.RifAddress].Value = val;
                    DataBase.SalvaModificheDB();
                }
                Target.Worksheet.Protect(Simboli.pwd);
            }
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
                if (!Workbook.fromErrorPane)
                {
                    DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.CheckOnly);
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
            Workbook.WB.SheetChange -= Handler.StoreEdit;

            int count = 0;
            foreach (RibbonToggleButton button in FrontOffice.Groups.First(g => g.Name == "groupAmbienti").Items)
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
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogModifica, "Attivato ambiente " + ambienteScelto.Name);
                DataBase.SwitchEnvironment(ambienteScelto.Name.Replace("btn", ""));
                btnAggiornaStruttura_Click(null, null);
            }

            Workbook.WB.SheetChange += Handler.StoreEdit;
            ambienteScelto.Checked = true;
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
            
            if(sender != null && e != null)
                response = System.Windows.Forms.MessageBox.Show("Eseguire l'aggiornamento della struttura?", Simboli.nomeApplicazione + " - ATTENZIONE!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
            
            if (response == System.Windows.Forms.DialogResult.Yes)
            {
                //se risposta positiva, mostro la splash screen
                SplashScreen.Show();

                Workbook.WB.SheetChange -= Handler.StoreEdit;
                Workbook.WB.Application.ScreenUpdating = false;
                Sheet.Proteggi(false);
                Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

                //verifico che la connessione sia funzionante
                if (DataBase.OpenConnection())
                {
                    //lancio l'aggiornamento della struttura
                    AggiornaStruttura();
                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogModifica, "Aggiorna struttura");
                }

                //refresh delle funzioni di check nel caso ci fossero
                _errorPane.RefreshCheck(_checkFunctions);

                Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                Sheet.Proteggi(true);
                Workbook.WB.Application.ScreenUpdating = true;
                Workbook.WB.SheetChange += Handler.StoreEdit;
                SplashScreen.Close();
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
            Workbook.WB.SheetChange -= Handler.StoreEdit;
            Workbook.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            DateTime dataOld = DataBase.DataAttiva;

            //apro il form calendario
            Forms.FormCalendar cal = new FormCalendar();
            object calDate = cal.ShowDialog();

            //verifico che la data sia stata cambiata
            if (calDate != null)
            {
                DateTime date = (DateTime)calDate;
                if (DataBase.OpenConnection())
                {
                    DataBase.RefreshAppSettings("DataInizio", date.ToString("yyyyMMdd"));
                    btnCalendar.Label = date.ToString("dddd dd MMM yyyy");

                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogModifica, "Cambio Data a " + btnCalendar.Label);

                    DataBase.RefreshDate(date);
                    DataBase.ConvertiParametriInformazioni();

                    DataView stato = DataBase.Select(DataBase.SP.CHECKMODIFICASTRUTTURA, "@DataOld=" + dataOld.ToString("yyyyMMdd") + ";@DataNew=" + date.ToString("yyyyMMdd")).DefaultView;

                    SplashScreen.Show();

                    if (stato.Count > 0 && stato[0]["Stato"].Equals(1))
                        AggiornaStruttura();
                    else
                        AggiornaDati();

                    Workbook.RefreshLog();
                    SplashScreen.Close();
                }
                else  //emergenza
                {
                    DataBase.RefreshAppSettings("DataInizio", date.ToString("yyyyMMdd"));
                    btnCalendar.Label = date.ToString("dddd dd MMM yyyy");
                    DataBase.RefreshDate(date);

                    foreach (Excel.Worksheet ws in Workbook.WB.Sheets)
                    {
                        if (ws.Name != "Log" && ws.Name != "Main")
                        {
                            Sheet s = new Sheet(ws);
                            s.AggiornaDateTitoli();
                        }
                    }

                    Riepilogo main = new Riepilogo(Workbook.WB.Sheets["Main"]);
                    main.RiepilogoInEmergenza();
                }

                _errorPane.RefreshCheck(_checkFunctions);
            }

            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Workbook.WB.Application.ScreenUpdating = true;
            Workbook.WB.SheetChange += Handler.StoreEdit;
        }
        /// <summary>
        /// Handler del click del tasto di selezione rampe. Apre il form per la selezione delle rampe ed esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnRampe_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.SheetChange -= Handler.StoreEdit;
            Workbook.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            //prendo il nome sheet e il range selezionato (per poter lavorare su più giorni nel caso ci fosse necessità)
            string sheet = Workbook.WB.ActiveSheet.Name;
            Excel.Range rng = Workbook.WB.Application.Selection;
            
            DefinedNames definedNames = new DefinedNames(sheet);
            FormSelezioneUP selUP = new FormSelezioneUP("PQNR_PROFILO");

            //controllo se nel range selezionato è definita un'entità
            if (sheet == "Iren Termo" && definedNames.IsDefined(rng.Row))
            {
                string nome = definedNames.GetNameByAddress(rng.Row, rng.Column);
                string siglaEntita = nome.Split(Simboli.UNION[0])[0];
                
                //controllo se l'entità ha la possibilità di selezionare le rampe
                DataView entitaInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
                entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaInformazione = 'PQNR_PROFILO'";

                if (entitaInformazioni.Count == 0)
                {
                    //avviso l'utente che l'entità selezionata non ha l'opzione
                    if (System.Windows.Forms.MessageBox.Show("L'operazione selezionata non è disponibile per l'UP selezionata, selezionarne un'altra dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes
                        && selUP.ShowDialog().ToString() != "")
                    {
                        Forms.FormRampe rampe = new FormRampe(Workbook.WB.Application.Selection);
                        rampe.ShowDialog();
                        rampe.Dispose();
                    }
                }
                else
                {
                    Forms.FormRampe rampe = new FormRampe(Workbook.WB.Application.Selection);
                    rampe.ShowDialog();
                    rampe.Dispose();
                }
            }
            //sono in un foglio diverso da Iren Termo o su una cella senza definizione di nomi
            else if (System.Windows.Forms.MessageBox.Show("Nessuna UP selezionata, selezionarne una dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes
                && selUP.ShowDialog().ToString() != "")
            {
                Forms.FormRampe rampe = new FormRampe(Workbook.WB.Application.Selection);
                rampe.ShowDialog();
                rampe.Dispose();
            }

            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Workbook.WB.Application.ScreenUpdating = true;
            Workbook.WB.SheetChange += Handler.StoreEdit;
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
                SplashScreen.Show();

                Workbook.WB.SheetChange -= Handler.StoreEdit;
                Workbook.WB.Application.ScreenUpdating = false;
                Sheet.Proteggi(false);
                Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

                if (DataBase.OpenConnection())
                {
                    AggiornaDati();
                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogModifica, "Aggiorna Dati");
                }

                _errorPane.RefreshCheck(_checkFunctions);

                SplashScreen.Close();

                Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                Sheet.Proteggi(true);
                Workbook.WB.Application.ScreenUpdating = true;
                Workbook.WB.SheetChange += Handler.StoreEdit;
            }
        }
        /// <summary>
        /// Handler del click del tasto delle azioni. Mostra il form delle azioni ed esegue il refresh dei check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAzioni_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            
            FormAzioni frmAz = new FormAzioni(new Esporta(), new Riepilogo(), new Carica());
            frmAz.ShowDialog();
            
            _errorPane.RefreshCheck(_checkFunctions);



            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Workbook.WB.Application.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di modifica. Attiva e disattiva la modifica foglio. Nel caso di disattivazione, aggiorna i check.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnModifica_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Simboli.ModificaDati = btnModifica.Checked;

            if (btnModifica.Checked) 
            {
                btnModifica.Image = Iren.ToolsExcel.Base.Properties.Resources.modificaSI_icon;
                btnModifica.Label = "Modifica SI";
            }
            else
            {
                _errorPane.RefreshCheck(_checkFunctions);

                //salva modifiche sul db
                Sheet.SalvaModifiche();
                DataBase.SalvaModificheDB();
                btnModifica.Image = Iren.ToolsExcel.Base.Properties.Resources.modificaNO_icon;
                btnModifica.Label = "Modifica NO";
            }
            Sheet.AbilitaModifica(btnModifica.Checked);

            Sheet.Proteggi(true);
            Workbook.WB.Application.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di Ottimizzazione.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnOttimizza_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.Application.ScreenUpdating = false;
            Workbook.WB.SheetChange -= Handler.StoreEdit;
            Sheet.Proteggi(false);

            Excel.Range rng = Workbook.WB.Application.Selection;

            DefinedNames definedNames = new DefinedNames(Workbook.WB.ActiveSheet.Name, DefinedNames.InitType.NamingOnly);

            //inizializzo ottimizzatore e il form di selezione entità per l'ottimo.
            Optimizer opt = new Optimizer();
            FormSelezioneUP selUP = new FormSelezioneUP("OTTIMO");

            if (definedNames.IsDefined(rng.Row))
            {
                string nome = definedNames.GetNameByAddress(rng.Row, rng.Column);
                string siglaEntita = nome.Split(Simboli.UNION[0])[0];

                DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
                
                if(categoriaEntita.Count > 0)
                    siglaEntita = categoriaEntita[0]["Gerarchia"] is DBNull ? siglaEntita : categoriaEntita[0]["Gerarchia"].ToString();

                DataView entitaInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
                entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaInformazione = 'OTTIMO'";

                if (entitaInformazioni.Count == 0)
                {
                    if(System.Windows.Forms.MessageBox.Show("L'UP selezionata non può essere ottimizzata, selezionarne un'altra dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes)
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
            Sheet.Proteggi(true);
            Workbook.WB.SheetChange += Handler.StoreEdit;
            Workbook.WB.Application.ScreenUpdating = true;
        }
        /// <summary>
        /// Handler del click del tasto di modifica parametri. Mostra il form di modifica dei parametri utente.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnConfigura_Click(object sender, RibbonControlEventArgs e)
        {
            FormConfig conf = new FormConfig();
            conf.ShowDialog();
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
            Simboli.EmergenzaForzata = btnForzaEmergenza.Checked;
        }
        /// <summary>
        /// Handler del click del tasto di chiusura. Chiude l'applicativo.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChiudi_Click(object sender, RibbonControlEventArgs e)
        {
            TextInfo ti = new CultureInfo("it-IT", false).TextInfo;
            var path = Utility.Workbook.GetUsrConfigElement("backup");
            string pathStr = Utility.ExportPath.PreparePath(path.Value);
            if (!Directory.Exists(pathStr))
                Directory.CreateDirectory(pathStr);

            string filename = ti.ToTitleCase(Simboli.nomeApplicazione).Replace(" ", "") + "_Backup_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsm"; 

            Globals.ThisWorkbook.SaveCopyAs(Path.Combine(pathStr, filename));
            Globals.ThisWorkbook.ThisApplication.Quit();
        }
        /// <summary>
        /// Handler del click del tasto per visualizzare l'actionsPane del documento.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnMostraErrorPane_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.ThisApplication.ScreenUpdating = false;
            Sheet.Proteggi(false);
            if (!Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane)
                Globals.ThisWorkbook.ThisApplication.DisplayDocumentActionTaskPane = true;
            
            _errorPane.SetDimension(Globals.ThisWorkbook.ActionsPane.Width, Globals.ThisWorkbook.ActionsPane.Height);
            
            //Check checkFunctions = new Check();
            //_errorPane.RefreshCheck(checkFunctions);
            Sheet.Proteggi(true);
            Globals.ThisWorkbook.ThisApplication.ScreenUpdating = true;
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Metodo di inizializzazione della Tab Front Office. Visualizza e abilita i tasti in base alle specifiche da DB. Da notare che se ci sono aggiornamenti, bisogna caricare la struttura e riavviare l'applicativo.
        /// </summary>
        private void Initialize()
        {
            _controls = new ControlCollection(this);
            DataView controlli = new DataView();
            
            if (DataBase.OpenConnection())
            {
                Struttura.CaricaApplicazioneRibbon();
                controlli = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE_RIBBON].DefaultView;
                DataBase.CloseConnection();
            }
            else
            {
                try
                {
                    controlli = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE_RIBBON].DefaultView;
                }
                catch
                {
                    controlli = null;
                }
            }

            if (controlli != null)
            {
                foreach (DataRowView controllo in controlli)
                {
                    Controls[controllo["NomeControllo"].ToString()].Visible = controllo["Visibile"].Equals("1");
                    Controls[controllo["NomeControllo"].ToString()].Enabled = controllo["Abilitato"].Equals("1");
                    if (controllo["Abilitato"].Equals("1"))
                        _enabledControls.Add(controllo["NomeControllo"].ToString());

                    if (Controls[controllo["NomeControllo"].ToString()].GetType().ToString().Contains("ToggleButton"))
                    {
                        ((RibbonToggleButton)Controls[controllo["NomeControllo"].ToString()]).Checked = controllo["Stato"].Equals("1");
                    }
                }

                List<RibbonGroup> groups = FrontOffice.Groups.ToList();
                foreach (RibbonGroup group in groups)
                    group.Visible = group.Items.Any(c => c.Visible);
            }
            else
            {
                foreach (RibbonControl control in Controls)
                {
                    control.Visible = true;
                    control.Enabled = true;
                    if (control.GetType().ToString().Contains("ToggleButton"))
                        ((RibbonToggleButton)control).Checked = false;
                }
            }
        }
        /// <summary>
        /// Metodo che seleziona il tasto corretto tra quelli degli applicativi presenti nella Tab Front Office. La selezione avviene in base all'ID applicazione scritto sul file di configurazione.
        /// </summary>
        private void CheckTastoApplicativo()
        {
            switch (Workbook.AppSettings("AppID"))
            {
                case "1":
                    btnOfferteMGP.Checked = true;
                    break;
                case "2":
                case "3":
                case "4":
                case "13":
                    btnInvioProgrammi.Checked = true;
                    break;
                case "5":
                    btnProgrammazioneImpianti.Checked = true;
                    break;
                case "6":
                    btnUnitCommitment.Checked = true;
                    break;
                case "7":
                    btnPrezziMSD.Checked = true;
                    break;
                case "8":
                    btnSistemaComandi.Checked = true;
                    break;
                case "9":
                    btnOfferteMSD.Checked = true;
                    break;
                case "10":
                    btnOfferteMB.Checked = true;
                    break;
                case "11":
                    btnValidazioneTL.Checked = true;
                    break;
                case "12":
                    btnPrevisioneCT.Checked = true;
                    break;
            }



        }
        /// <summary>
        /// Disabilito tutti i tasti nel caso in cui, ad esempio in seguito a un rilascio, il foglio parta completamente da 0. Disabilita tutti i tasti eccetto Aggiorna Struttura che consente all'utente di rendere operativo il foglio.
        /// </summary>
        private void DisabilitaTasti()
        {
            foreach (RibbonControl control in Controls)
            {
                if(control.Name != "btnAggiornaStruttura")
                    control.Enabled = false;
            }
            _allDisabled = true;
        }
        /// <summary>
        /// Abilito tutti i tasti nel caso in cui, ad esempio in seguito a un rilascio, questi vengano disabilitati da DisabilitaTasti.
        /// </summary>
        private void AbilitaTasti()
        {
            foreach (string control in _enabledControls)
                Controls[control].Enabled = true;

            _allDisabled = false;
        }
        /// <summary>
        /// Attiva l'aggiornamento della struttura del foglio che consiste in:
        ///  - azzerare il dataset locale 
        ///  - caricarlo nuovamente dal DB 
        ///  - generare i fogli che non esistono
        ///  - lanciare la routine per ri-creare la struttura
        ///  - caricare la struttura del riepilogo.
        /// </summary>
        private void AggiornaStruttura()
        {
            SplashScreen.UpdateStatus("Carico struttura dal DB");
            Struttura.AggiornaStrutturaDati();

            DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "Operativa = 1";

            foreach (DataRowView categoria in categorie)
            {
                Excel.Worksheet ws;
                try
                {
                    ws = Workbook.WB.Worksheets[categoria["DesCategoria"].ToString()];
                }
                catch
                {
                    ws = (Excel.Worksheet)Workbook.WB.Worksheets.Add(Workbook.WB.Worksheets["Log"]);
                    ws.Name = categoria["DesCategoria"].ToString();
                    ws.Select();
                    Workbook.WB.Application.Windows[1].DisplayGridlines = false;
#if !DEBUG
                    Workbook.WB.Application.ActiveWindow.DisplayHeadings = false;
#endif
                }
            }

            Workbook.WB.Sheets["Main"].Select();
            Riepilogo main = new Riepilogo(Workbook.WB.Sheets["Main"]);
            SplashScreen.UpdateStatus("Aggiorno struttura Riepilogo");
            main.LoadStructure();            

            foreach (Excel.Worksheet ws in Workbook.WB.Sheets)
            {
                if (ws.Name != "Log" && ws.Name != "Main")
                {
                    Sheet s = new Sheet(ws);
                    SplashScreen.UpdateStatus("Aggiorno struttura " + ws.Name);
                    s.LoadStructure();
                }
            }

            SplashScreen.UpdateStatus("Salvo struttura in locale");
            Workbook.DumpDataSet();

            Globals.Main.Select();
            //Workbook.WB.Sheets["Main"].Select();
            Globals.Main.Range["A1"].Select();
            //Workbook.WB.ActiveSheet.Range["A1"].Select();
            Workbook.WB.Application.WindowState = Excel.XlWindowState.xlMaximized;

            if (_allDisabled)
                AbilitaTasti();
        }
        /// <summary>
        /// Attiva l'aggiornamento dei dati contenuti nel foglio senza però alterare la struttura del foglio stesso.
        /// </summary>
        private void AggiornaDati()
        {
            foreach (Excel.Worksheet ws in Workbook.WB.Sheets)
            {
                if (ws.Name != "Log" && ws.Name != "Main")
                {
                    Sheet s = new Sheet(ws);
                    SplashScreen.UpdateStatus("Aggiornamento dati " + ws.Name);
                    s.UpdateData(true);
                }
            }
            Riepilogo main = new Riepilogo(Workbook.WB.Sheets["Main"]);
            SplashScreen.UpdateStatus("Aggiornamento Riepilogo");
            main.UpdateRiepilogo();
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

        private ToolsExcelRibbon _ribbon;
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

        internal ControlCollection(ToolsExcelRibbon ribbon)
        {
            _ribbon = ribbon;
            List<RibbonGroup> groups = ribbon.FrontOffice.Groups.ToList();

            foreach (RibbonGroup group in groups)
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

        private ToolsExcelRibbon _ribbon;
        private int _pos = -1;
        private int _max = -1;

        #endregion

        #region Costruttori

        public ControlEnumerator(ToolsExcelRibbon ribbon)
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
