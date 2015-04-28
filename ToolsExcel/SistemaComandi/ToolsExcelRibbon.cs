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

// ***************************************************** SISTEMA COMANDI ***************************************************** //

namespace Iren.ToolsExcel
{
    public partial class ToolsExcelRibbon
    {
        #region Variabili        
        
        #endregion

        #region Eventi

        private void ToolsExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            this.RibbonUI.ActivateTab(FrontOffice.ControlId.CustomId);

            if (Workbook.WB.Sheets.Count <= 2)
                AbilitaTasti(false);

            CheckTastoApplicativo();

            DateTime cfgDate = DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], "yyyyMMdd", CultureInfo.InvariantCulture);
            btnCalendar.Label = cfgDate.ToString("dddd dd MMM yyyy");

            btnModifica.Checked = false;

            
            //configuro gli ambienti selezionabili
            string[] ambienti = ConfigurationManager.AppSettings["AmbientiVisibili"].Split('|');
            foreach (string ambiente in ambienti)
                groupAmbienti.Items.OfType<RibbonToggleButton>().Where(btn => btn.Name == "btn" + ambiente).ToArray()[0].Visible = true;

            //seleziono l'ambiente attivo
            groupAmbienti.Items.OfType<RibbonToggleButton>().Where(btn => btn.Name == "btn" + ConfigurationManager.AppSettings["DB"]).ToArray()[0].Checked = true;

            //configuro i tasti visibili
            if (ConfigurationManager.AppSettings["RampeVisible"] != null && ConfigurationManager.AppSettings["RampeVisible"].ToLowerInvariant() == "false")
                btnRampe.Visible = false;

            //se esce con qualche errore il tasto mantiene lo stato a cui era impostato
            btnModifica.Checked = false;
            btnModifica.Image = Iren.ToolsExcel.Base.Properties.Resources.modifica_no_icon;
            btnModifica.Label = "Modifica NO";

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
                AbilitaTasti(true);
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
                            Struttura.AggiornaStrutturaDati();
                            AggiornaStruttura();
                        }
                        else
                        {
                            AggiornaDati();
                        }

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

        private void AbilitaTasti(bool abilita)
        {
            btnCalendar.Enabled = abilita;
            btnAzioni.Enabled = abilita;
            btnRampe.Enabled = abilita;
            btnAggiornaDati.Enabled = abilita;
            btnModifica.Enabled = abilita;
            btnOttimizza.Enabled = abilita;
            btnCalendar.Enabled = abilita;
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
            //SplashForm.LabelText = "Salvo struttura in locale";
            Workbook.DumpDataSet();
            
            Workbook.WB.Sheets["Main"].Select();
            Workbook.WB.Application.WindowState = Excel.XlWindowState.xlMaximized;
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
    }
}
