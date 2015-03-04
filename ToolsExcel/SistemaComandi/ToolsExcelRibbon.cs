using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Core;
using Iren.ToolsExcel.Forms;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.ToolsExcel
{
    public partial class ToolsExcelRibbon
    {
        #region Variabili

        LoaderScreen loader = new LoaderScreen();
        
        #endregion

        #region Eventi

        private void ToolsExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {      
            if (UtilityWB.WB.Sheets.Count <= 2)
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

        }
        private void btnSelezionaAmbiente_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton ambienteScelto = (RibbonToggleButton)sender;
            UtilityWB.WB.SheetChange -= Handler.StoreEdit;

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
                //TODO riabilitare log!!
                //UtilityWB.InsertLog(DataBase.TipologiaLOG.LogModifica, "Attivato ambiente " + ambienteScelto.Name);
                UtilityDB.SwitchEnvironment(ambienteScelto.Name.Replace("btn", ""));
                btnAggiornaStruttura_Click(null, null);
            }

            UtilityWB.WB.SheetChange += Handler.StoreEdit;
            ambienteScelto.Checked = true;
        }
        private void btnAggiornaStruttura_Click(object sender, RibbonControlEventArgs e)
        {
            UtilityWB.WB.SheetChange -= Handler.StoreEdit;
            UtilityWB.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            if (UtilityDB.DB.OpenConnection())
            {
                AggiornaStruttura();
                //TODO riabilitare log!!
                //UtilityWB.InsertLog(DataBase.TipologiaLOG.LogModifica, "Aggiorna struttura");

                UtilityDB.DB.CloseConnection();
            }

            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            UtilityWB.WB.Application.ScreenUpdating = true;
            UtilityWB.WB.SheetChange += Handler.StoreEdit;

            AbilitaTasti(true);
        }
        private void btnCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            UtilityWB.WB.SheetChange -= Handler.StoreEdit;
            UtilityWB.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            Forms.FormCalendar cal = new FormCalendar();
            cal.ShowDialog();

            DateTime dataOld = DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], "yyyyMMdd", CultureInfo.InvariantCulture);

            if (cal.Date != null)
            {
                if (dataOld != cal.Date.Value)
                {
                    if (UtilityDB.DB.OpenConnection())
                    {
                        UtilityDB.RefreshAppSettings("DataInizio", cal.Date.Value.ToString("yyyyMMdd"));
                        btnCalendar.Label = cal.Date.Value.ToString("dddd dd MMM yyyy");

                        //TODO riabilitare log!!
                        //UtilityWB.InsertLog(DataBase.TipologiaLOG.LogModifica, "Cambio Data a " + btnCalendar.Label);

                        UtilityDB.RefreshDate(cal.Date.Value);
                        UtilityDB.ConvertiParametriInformazioni();

                        DataView stato = UtilityDB.DB.Select(UtilityDB.SP.CHECKMODIFICASTRUTTURA, "@DataOld=" + dataOld.ToString("yyyyMMdd") + ";@DataNew=" + cal.Date.Value.ToString("yyyyMMdd")).DefaultView;
                        if (stato.Count > 0 && stato[0]["Stato"].Equals("1"))
                        {
                            UtilityStruttura.AggiornaStrutturaDati();
                            AggiornaStruttura();
                        }
                        else
                        {
                            AggiornaDati(all: true);
                        }
                        UtilityDB.DB.CloseConnection();
                    }
                    else  //emergenza
                    {
                        UtilityDB.RefreshAppSettings("DataInizio", cal.Date.Value.ToString("yyyyMMdd"));
                        btnCalendar.Label = cal.Date.Value.ToString("dddd dd MMM yyyy");
                        UtilityDB.RefreshDate(cal.Date.Value);

                        foreach (Excel.Worksheet ws in UtilityWB.WB.Sheets)
                        {
                            if (ws.Name != "Log" && ws.Name != "Main")
                            {
                                Sheet s = new Sheet(ws);
                                s.AggiornaDateTitoli();
                            }
                        }

                        Riepilogo main = new Riepilogo(UtilityWB.WB.Sheets["Main"]);
                        main.RiepilogoInEmergenza();
                    }
                }
            }
            cal.Dispose();

            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            UtilityWB.WB.Application.ScreenUpdating = true;
            UtilityWB.WB.SheetChange += Handler.StoreEdit;
        }
        private void btnRampe_Click(object sender, RibbonControlEventArgs e)
        {
            UtilityWB.WB.SheetChange -= Handler.StoreEdit;
            UtilityWB.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            Excel.Worksheet ws = (Excel.Worksheet)UtilityWB.WB.ActiveSheet;
            Excel.Range rng = UtilityWB.WB.Application.Selection;
            
            DefinedNames nomiDefiniti = new DefinedNames(ws.Name);

            string siglaEntita = "";

            if (ws.Name == "Iren Termo" && nomiDefiniti.IsDefined(rng.Row, rng.Column))
            {
                string nome = nomiDefiniti[rng.Row, rng.Column][0];
                siglaEntita = nome.Split(char.Parse(Simboli.UNION))[0];

                DataView entitaInformazioni = UtilityDB.LocalDB.Tables[UtilityDB.Tab.ENTITAINFORMAZIONE].DefaultView;
                entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaInformazione = 'PQNR_PROFILO'";

                if (entitaInformazioni.Count == 0 
                    && System.Windows.Forms.MessageBox.Show("L'UP selezionata non può essere ottimizzata, selezionarne un'altra dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes
                    && SelezionaUP("PQNR_PROFILO", out siglaEntita, out nomiDefiniti, out rng))
                {
                        Forms.FormRampe rampe = new FormRampe(nomiDefiniti, rng);
                        rampe.ShowDialog();
                        rampe.Dispose();
                }
                else
                {
                    Forms.FormRampe rampe = new FormRampe(nomiDefiniti, rng);
                    rampe.ShowDialog();
                    rampe.Dispose();
                }
            }
            else if (System.Windows.Forms.MessageBox.Show("Nessuna UP selezionata, selezionarne una dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes
                && SelezionaUP("PQNR_PROFILO", out siglaEntita, out nomiDefiniti, out rng))
            {
                Forms.FormRampe rampe = new FormRampe(nomiDefiniti, rng);
                rampe.ShowDialog();
                rampe.Dispose();
            }

            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            UtilityWB.WB.Application.ScreenUpdating = true;
            UtilityWB.WB.SheetChange += Handler.StoreEdit;
        }
        private void btnAggiornaDati_Click(object sender, RibbonControlEventArgs e)
        {
            UtilityWB.WB.SheetChange -= Handler.StoreEdit;
            UtilityWB.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            if (UtilityDB.DB.OpenConnection())
            {
                AggiornaDati(all: false);

                //TODO riabilitare log!!
                //UtilityWB.InsertLog(DataBase.TipologiaLOG.LogModifica, "Aggiorna Dati");

                UtilityDB.DB.CloseConnection();
            }
            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            UtilityWB.WB.Application.ScreenUpdating = true;
            UtilityWB.WB.SheetChange += Handler.StoreEdit;
        }
        private void btnAzioni_Click(object sender, RibbonControlEventArgs e)
        {
            UtilityWB.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            CustFormAzioni frmAz = new CustFormAzioni();
            frmAz.ShowDialog();

            UtilityWB.WB.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            UtilityWB.WB.Application.ScreenUpdating = true;
        }
        private void btnModifica_Click(object sender, RibbonControlEventArgs e)
        {
            UtilityWB.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Simboli.ModificaDati = btnModifica.Checked;

            Sheet.AbilitaModifica(btnModifica.Checked);
            if (btnModifica.Checked) 
                btnModifica.Image = global::Iren.ToolsExcel.Base.Properties.Resources.edit_validated_icon;
            else
            {
                //Salva modifiche su db
                UtilityDB.SalvaModificheDB();
                btnModifica.Image = global::Iren.ToolsExcel.Base.Properties.Resources.edit_not_validated_icon;
            }
            Sheet.Proteggi(true);
            UtilityWB.WB.Application.ScreenUpdating = true;
        }
        private void btnOttimizza_Click(object sender, RibbonControlEventArgs e)
        {
            UtilityWB.WB.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);

            Excel.Worksheet ws = (Excel.Worksheet)UtilityWB.WB.ActiveSheet;
            Excel.Range rng = UtilityWB.WB.Application.Selection;

            DefinedNames nomiDefiniti = new DefinedNames(ws.Name);

            string siglaEntita = "";


            if (nomiDefiniti.IsDefined(rng.Row, rng.Column))
            {
                siglaEntita = nomiDefiniti[rng.Row, rng.Column][0].Split(char.Parse(Simboli.UNION))[0];

                DataView entitaInformazioni = UtilityDB.LocalDB.Tables[UtilityDB.Tab.ENTITAINFORMAZIONE].DefaultView;
                entitaInformazioni.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaInformazione = 'OTTIMO'";

                if (entitaInformazioni.Count == 0 
                    && System.Windows.Forms.MessageBox.Show("L'UP selezionata non può essere ottimizzata, selezionarne un'altra dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes 
                    && SelezionaUP("OTTIMO", out siglaEntita, out nomiDefiniti, out rng))
                {
                    Optimizer.EseguiOttimizzazione(siglaEntita);
                }
                else
                {
                    Optimizer.EseguiOttimizzazione(siglaEntita);
                }
            }
            else if (System.Windows.Forms.MessageBox.Show("Nessuna UP selezionata, selezionarne una dall'elenco?", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Exclamation) == System.Windows.Forms.DialogResult.Yes
                    && SelezionaUP("OTTIMO", out siglaEntita, out nomiDefiniti, out rng))
            {
                Optimizer.EseguiOttimizzazione(siglaEntita);
            }

            Sheet.Proteggi(true);
            UtilityWB.WB.Application.ScreenUpdating = true;
        }
        private void btnConfigura_Click(object sender, RibbonControlEventArgs e)
        {
            FormConfig conf = new FormConfig();
            conf.ShowDialog();
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

        private bool SelezionaUP(string siglaInformazione, out string siglaEntita, out DefinedNames nomiDefiniti, out Excel.Range rng)
        {
            FormSelezioneUP selUP = new FormSelezioneUP(siglaInformazione);

            selUP.ShowDialog();

            nomiDefiniti = null;
            rng = null;
            siglaEntita = "";

            if (!selUP.IsCanceld && selUP.HasSelection)
            {
                siglaEntita = selUP.SiglaEntita;
                string nome = DefinedNames.GetName(selUP.SiglaEntita, "T", "DATA1");
                string foglio = DefinedNames.GetSheetName(nome);
                nomiDefiniti = new DefinedNames(foglio);
                Tuple<int, int>[] celle = nomiDefiniti.GetRange(nome);

                Excel.Worksheet ws = UtilityWB.WB.Application.Sheets[foglio];
                ((Excel._Worksheet)ws).Activate();
                rng = ws.Range[ws.Cells[celle[0].Item1, celle[0].Item2], ws.Cells[celle[1].Item1, celle[1].Item2]];
                rng.Select();
                UtilityWB.WB.Application.ActiveWindow.SmallScroll(celle[0].Item1 - ws.Application.ActiveWindow.VisibleRange.Cells[1, 1].Row - 1);
            }
            selUP.Dispose();
            return !selUP.IsCanceld && selUP.HasSelection;
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
            UtilityStruttura.AggiornaStrutturaDati();

            DataView categorie = UtilityDB.LocalDB.Tables[UtilityDB.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "Operativa = 1";

            foreach (DataRowView categoria in categorie)
            {
                Excel.Worksheet ws;
                try
                {
                    ws = UtilityWB.WB.Worksheets[categoria["DesCategoria"].ToString()];
                }
                catch
                {
                    ws = (Excel.Worksheet)UtilityWB.WB.Worksheets.Add(UtilityWB.WB.Worksheets["Log"]);
                    ws.Name = categoria["DesCategoria"].ToString();
                    ws.Select();
                    UtilityWB.WB.Application.Windows[1].DisplayGridlines = false;                    
                }
            }

            Riepilogo main = new Riepilogo(UtilityWB.WB.Sheets["Main"]);
            main.LoadStructure();

            foreach (Excel.Worksheet ws in UtilityWB.WB.Sheets)
            {
                if (ws.Name != "Log" && ws.Name != "Main")
                {
                    Sheet s = new Sheet(ws);
                    s.LoadStructure();
                }
            }

            UtilityWB.DumpDataSet();
            
            UtilityWB.WB.Sheets["Main"].Select();
            UtilityWB.WB.Application.WindowState = Excel.XlWindowState.xlMaximized;
        }
        private void AggiornaDati(bool all)
        {
            foreach (Excel.Worksheet ws in UtilityWB.WB.Sheets)
            {
                if (ws.Name != "Log" && ws.Name != "Main")
                {
                    Sheet s = new Sheet(ws);
                    s.UpdateData(all);
                }
            }
            if (all)
            {
                Riepilogo main = new Riepilogo(UtilityWB.WB.Sheets["Main"]);
                main.UpdateRiepilogo();
            }

            //Log
            //CommonFunctions.InitLog();
        }

        #endregion

        private void btnProgrammi_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton btn = (RibbonToggleButton)sender;

            if (!btn.Checked)
            {
                btn.Checked = true;
            }
            else
            {
                //TODO aprire gli altri file!!!!!!
                switch (btn.Name)
                {
                    case "btnInvioProgrammi":
                        break;
                }


            }
        }

    }
}
