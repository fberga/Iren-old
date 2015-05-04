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

// ***************************************************** SISTEMA COMANDI ***************************************************** //

namespace Iren.ToolsExcel
{
    public partial class ToolsExcelRibbon
    {
        #region Variabili
        
        private ControlCollection _controls;
        private List<string> _enabledControls = new List<string>();
        private bool _allDisabled = false;

        #endregion

        #region Proprietà

        public ControlCollection Controls
        {
            get { return _controls; }
        }

        #endregion

        #region Eventi

        private void ToolsExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            Initialize();

            if (Workbook.WB.Sheets.Count <= 2)
                DisabilitaTasti();

            this.RibbonUI.ActivateTab(FrontOffice.ControlId.CustomId);

            DateTime cfgDate = DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], "yyyyMMdd", CultureInfo.InvariantCulture);
            btnCalendar.Label = cfgDate.ToString("dddd dd MMM yyyy");

            //seleziono l'ambiente attivo
            ((RibbonToggleButton)Controls["btn" + ConfigurationManager.AppSettings["DB"]]).Checked = true;

            //se esce con qualche errore il tasto mantiene lo stato a cui era impostato
            btnModifica.Checked = false;
            btnModifica.Image = Iren.ToolsExcel.Base.Properties.Resources.modifica_no_icon;
            btnModifica.Label = "Modifica NO";
            try
            {
                Sheet.AbilitaModifica(false);
            }
            catch 
            { }
        }
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
        private void btnAggiornaStruttura_Click(object sender, RibbonControlEventArgs e)
        {
            var response = System.Windows.Forms.MessageBox.Show("Eseguire l'aggiornamento della struttura?", Simboli.nomeApplicazione + " - ATTENZIONE!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
            if (response == System.Windows.Forms.DialogResult.Yes)
            {
                SplashScreen.Show();

                Workbook.WB.SheetChange -= Handler.StoreEdit;
                Workbook.WB.Application.ScreenUpdating = false;
                Sheet.Proteggi(false);
                Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

                if (DataBase.OpenConnection())
                {
                    AggiornaStruttura();
                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogModifica, "Aggiorna struttura");
                    DataBase.DB.CloseConnection();
                }

                Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                Sheet.Proteggi(true);
                Workbook.WB.Application.ScreenUpdating = true;
                Workbook.WB.SheetChange += Handler.StoreEdit;
                SplashScreen.Close();
            }
        }
        private void btnCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.SheetChange -= Handler.StoreEdit;
            Workbook.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            Forms.FormCalendar cal = new FormCalendar();
            cal.ShowDialog();

            DateTime dataOld = DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], "yyyyMMdd", CultureInfo.InvariantCulture);

            if (cal.Date != null)
            {
                if (dataOld != cal.Date.Value)
                {
                    if (DataBase.OpenConnection())
                    {
                        DataBase.RefreshAppSettings("DataInizio", cal.Date.Value.ToString("yyyyMMdd"));
                        btnCalendar.Label = cal.Date.Value.ToString("dddd dd MMM yyyy");

                        Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogModifica, "Cambio Data a " + btnCalendar.Label);

                        DataBase.RefreshDate(cal.Date.Value);
                        DataBase.ConvertiParametriInformazioni();

                        DataView stato = DataBase.DB.Select(DataBase.SP.CHECKMODIFICASTRUTTURA, "@DataOld=" + dataOld.ToString("yyyyMMdd") + ";@DataNew=" + cal.Date.Value.ToString("yyyyMMdd")).DefaultView;

                        SplashScreen.Show();

                        if (stato.Count > 0 && stato[0]["Stato"].Equals(1))
                        {
                            //Struttura.AggiornaStrutturaDati();
                            AggiornaStruttura();
                        }
                        else
                        {
                            AggiornaDati();
                        }

                        Workbook.RefreshLog();
                        SplashScreen.Close();
                    }
                    else  //emergenza
                    {
                        DataBase.RefreshAppSettings("DataInizio", cal.Date.Value.ToString("yyyyMMdd"));
                        btnCalendar.Label = cal.Date.Value.ToString("dddd dd MMM yyyy");
                        DataBase.RefreshDate(cal.Date.Value);

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
                }
            }
            cal.Dispose();

            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Workbook.WB.Application.ScreenUpdating = true;
            Workbook.WB.SheetChange += Handler.StoreEdit;
        }
        private void btnRampe_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.SheetChange -= Handler.StoreEdit;
            Workbook.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            string sheet = Workbook.WB.ActiveSheet.Name;
            Excel.Range rng = Workbook.WB.Application.Selection;
            
            NewDefinedNames newNomiDefiniti = new NewDefinedNames(sheet, NewDefinedNames.InitType.NamingOnly);
            FormSelezioneUP selUP = new FormSelezioneUP("PQNR_PROFILO");

            if (sheet == "Iren Termo" && newNomiDefiniti.IsDefined(rng.Row))
            {
                string nome = newNomiDefiniti.GetNameByAddress(rng.Row, rng.Column);
                string siglaEntita = nome.Split(Simboli.UNION[0])[0];

                DataView entitaInformazioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_INFORMAZIONE].DefaultView;
                entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaInformazione = 'PQNR_PROFILO'";

                if (entitaInformazioni.Count == 0)
                {
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
                    DataBase.DB.CloseConnection();
                }

                SplashScreen.Close();

                Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                Sheet.Proteggi(true);
                Workbook.WB.Application.ScreenUpdating = true;
                Workbook.WB.SheetChange += Handler.StoreEdit;
            }
        }
        private void btnAzioni_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            
            FormAzioni frmAz = new FormAzioni(new Esporta(), new Riepilogo());
            frmAz.ShowDialog();

            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Workbook.WB.Application.ScreenUpdating = true;
        }
        private void btnModifica_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Simboli.ModificaDati = btnModifica.Checked;

            Sheet.AbilitaModifica(btnModifica.Checked);
            if (btnModifica.Checked) 
            {
                btnModifica.Image = Iren.ToolsExcel.Base.Properties.Resources.modifica_icon;
                btnModifica.Label = "Modifica SI";
            }
            else
            {
                //Salva modifiche su db
                Sheet.SalvaModifiche();
                DataBase.SalvaModificheDB();
                btnModifica.Image = Iren.ToolsExcel.Base.Properties.Resources.modifica_no_icon;
                btnModifica.Label = "Modifica NO";
            }
            Sheet.Proteggi(true);
            Workbook.WB.Application.ScreenUpdating = true;
        }
        private void btnOttimizza_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.Application.ScreenUpdating = false;
            Workbook.WB.SheetChange -= Handler.StoreEdit;
            Sheet.Proteggi(false);

            Excel.Range rng = Workbook.WB.Application.Selection;

            NewDefinedNames newNomiDefiniti = new NewDefinedNames(Workbook.WB.ActiveSheet.Name, NewDefinedNames.InitType.NamingOnly);

            Optimizer opt = new Optimizer();
            FormSelezioneUP selUP = new FormSelezioneUP("OTTIMO");

            if (newNomiDefiniti.IsDefined(rng.Row))
            {
                string nome = newNomiDefiniti.GetNameByAddress(rng.Row, rng.Column);
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
                            SplashScreen.UpdateStatus("Ottimizzazione " + siglaEntita + " in corso...");
                            opt.EseguiOttimizzazione(siglaEntita);
                            SplashScreen.Close();
                        }
                    }
                }
                else
                {
                    SplashScreen.Show();
                    SplashScreen.UpdateStatus("Ottimizzazione " + siglaEntita + " in corso...");
                    opt.EseguiOttimizzazione(siglaEntita);
                    SplashScreen.Close();
                }
            }
            else if (System.Windows.Forms.MessageBox.Show("Nessuna UP selezionata, selezionarne una dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes)
            {
                object siglaEntita = selUP.ShowDialog();
                if (siglaEntita != null)
                {
                    SplashScreen.Show();
                    SplashScreen.UpdateStatus("Ottimizzazione " + siglaEntita + " in corso...");
                    opt.EseguiOttimizzazione(siglaEntita);
                    SplashScreen.Close();
                }
            }

            Sheet.Proteggi(true);
            Workbook.WB.SheetChange += Handler.StoreEdit;
            Workbook.WB.Application.ScreenUpdating = true;
        }
        private void btnConfigura_Click(object sender, RibbonControlEventArgs e)
        {
            FormConfig conf = new FormConfig();
            conf.ShowDialog();
        }
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
        private void btnForzaEmergenza_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook.WB.SheetChange -= Handler.StoreEdit;
            Workbook.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            Simboli.EmergenzaForzata = btnForzaEmergenza.Checked;

            Riepilogo main = new Riepilogo(Workbook.WB.Sheets["Main"]);
            if (btnForzaEmergenza.Checked)
            {
                main.RiepilogoInEmergenza();
            }
            else
            {
                if (DataBase.OpenConnection())
                {
                    main.UpdateRiepilogo();
                    DataBase.DB.CloseConnection();
                }
            }

            Workbook.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Workbook.WB.Application.ScreenUpdating = true;
            Workbook.WB.SheetChange += Handler.StoreEdit;
        }
        #endregion

        #region Metodi



        private void Initialize()
        {
            _controls = new ControlCollection(this);
            DataView controlli = new DataView();
            try
            {
                controlli = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE_RIBBON].DefaultView;
            }
            catch 
            {
                if (DataBase.OpenConnection())
                {
                    Struttura.CaricaApplicazioneRibbon();
                    controlli = DataBase.LocalDB.Tables[DataBase.Tab.APPLICAZIONE_RIBBON].DefaultView;
                    DataBase.CloseConnection();
                }
            }

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

        }

        private void CheckTastoApplicativo()
        {
            switch (ConfigurationManager.AppSettings["AppID"])
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

        private void DisabilitaTasti()
        {
            foreach (RibbonControl control in Controls)
            {
                if(control.Name != "btnAggiornaStruttura")
                    control.Enabled = false;
            }
            _allDisabled = true;
        }
        private void AbilitaTasti()
        {
            foreach (string control in _enabledControls)
                Controls[control].Enabled = true;

            _allDisabled = false;
        }

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
            
            Workbook.WB.Sheets["Main"].Select();
            Workbook.WB.Application.WindowState = Excel.XlWindowState.xlMaximized;

            if (_allDisabled)
                AbilitaTasti();
        }
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

        private void btnChiudi_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.ThisApplication.Quit();
        }
    }

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

        #endregion
    }
    public class ControlEnumerator : IEnumerator
    {
        ToolsExcelRibbon _ribbon;
        int _pos = -1;
        int _max = -1;

        public ControlEnumerator(ToolsExcelRibbon ribbon)
        {
            _ribbon = ribbon;
            _max = ribbon.Controls.Count;
        }

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
    }
}
