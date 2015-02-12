using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Configuration;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using Iren.FrontOffice.Core;
using Iren.FrontOffice.Forms;
using Iren.FrontOffice.Base;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Iren.FrontOffice.Tools
{
    public partial class ToolsExcelRibbon
    {
        #region Variabili

        LoaderScreen loader = new LoaderScreen();
        
        #endregion

        #region Eventi

        private void ToolsExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            if (Globals.ThisWorkbook.Sheets.Count <= 2)
            {//disabilito tutti i tasti
                AbilitaTasti(false);
            }

            DateTime cfgDate = DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], "yyyyMMdd", CultureInfo.InvariantCulture);
            btnCalendar.Label = cfgDate.ToString("dddd dd MMM yyyy");

            btnModifica.Checked = false;

            string[] ambienti = ConfigurationManager.AppSettings["AmbientiVisibili"].Split('|');

            foreach (string ambiente in ambienti)
            {
                groupAmbienti.Items.OfType<RibbonToggleButton>().Where(btn => btn.Name == ambiente).ToArray()[0].Visible = true;
            }

            groupAmbienti.Items.OfType<RibbonToggleButton>().Where(btn => btn.Name == ConfigurationManager.AppSettings["DB"]).ToArray()[0].Checked = true;
        }

        private void btnSelezionaAmbiente_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonToggleButton ambienteScelto = (RibbonToggleButton)sender;

            int count = 0;
            foreach (RibbonToggleButton button in FrontOffice.Groups.Last().Items)
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
                //CommonFunctions.InsertLog(DataBase.TipologiaLOG.LogModifica, "Attivato ambiente " + ambienteScelto.Name);
                CommonFunctions.SwitchEnvironment(ambienteScelto.Name);
                btnAggiornaStruttura_Click(null, null);
            }

            ambienteScelto.Checked = true;
        }
        private void btnAggiornaStruttura_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.ThisApplication.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            if (CommonFunctions.DB.OpenConnection())
            {
                AggiornaStruttura();
                //TODO riabilitare log!!
                //CommonFunctions.InsertLog(DataBase.TipologiaLOG.LogModifica, "Aggiorna struttura");

                CommonFunctions.DB.CloseConnection();
            }

            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Globals.ThisWorkbook.ThisApplication.ScreenUpdating = true;

            AbilitaTasti(true);
        }
        private void btnCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            Forms.frmCALENDAR cal = new frmCALENDAR();
            cal.Text = Simboli.nomeApplicazione;
            cal.ShowDialog();


            DateTime dataOld = DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], "yyyyMMdd", CultureInfo.InvariantCulture);

            if (cal.Date != null)
            {
                if (dataOld != cal.Date.Value)
                {
                    if (CommonFunctions.DB.OpenConnection() && CommonFunctions.DB.StatoDB()[DataBase.NomiDB.SQLSERVER] == ConnectionState.Open)
                    {

                        CommonFunctions.RefreshAppSettings("DataInizio", cal.Date.Value.ToString("yyyyMMdd"));

                        btnCalendar.Label = cal.Date.Value.ToString("dddd dd MMM yyyy");

                        //TODO riabilitare log!!
                        //CommonFunctions.InsertLog(DataBase.TipologiaLOG.LogModifica, "Cambio Data a " + btnCalendar.Label);

                        CommonFunctions.RefreshDate(cal.Date.Value);
                        CommonFunctions.ConvertiParametriInformazioni();

                        DataView stato = CommonFunctions.DB.Select("spCheckModificaStruttura", "@DataOld=" + dataOld.ToString("yyyyMMdd") + ";@DataNew=" + cal.Date.Value.ToString("yyyyMMdd")).DefaultView;
                        if (stato.Count > 0 && stato[0]["Stato"].Equals("1"))
                        {
                            CommonFunctions.AggiornaStrutturaDati();
                            AggiornaStruttura();
                        }
                        else
                        {
                            AggiornaDati();
                        }
                        CommonFunctions.DB.CloseConnection();
                    }
                }
            }
            cal.Dispose();

            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }
        private void btnRampe_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            Excel.Worksheet ws = (Excel.Worksheet)Globals.ThisWorkbook.ActiveSheet;

            if (ws.Name == "Iren Termo")
            {
                Excel.Range rng = Globals.ThisWorkbook.Application.Selection;
                DefinedNames nomiDefiniti = new DefinedNames(ws.Name);

                string[] nome = nomiDefiniti[rng.Row, rng.Column];

                if (nome != null)
                {
                    string up = nome[0].Split(Simboli.UNION[0])[0];

                    string suffissoData = Regex.Match(nome[0], @"DATA\d+").Value;
                    suffissoData = suffissoData == "" ? "DATA1" : suffissoData;

                    DataView proprieta = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAPROPRIETA].DefaultView;
                    proprieta.RowFilter = "SiglaEntita = '" + up + "' AND SiglaProprieta = 'SISTEMA_COMANDI_PRIF'";
                    double pRif = 0;
                    if (proprieta.Count > 0)
                        pRif = Double.Parse(proprieta[0]["Valore"].ToString());

                    DataView categoriaEntita = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIAENTITA].DefaultView;
                    categoriaEntita.RowFilter = "SiglaEntita = '" + up + "'";
                    string desEntita = categoriaEntita[0]["DesEntita"].ToString();

                    DataView entitaRampa = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITARAMPA].DefaultView;
                    entitaRampa.RowFilter = "SiglaEntita = '" + up + "'";
                    object[] sigleRampa = entitaRampa.ToTable(false, "SiglaRampa").AsEnumerable().Select(r => r["SiglaRampa"]).ToArray();

                    Tuple<int, int>[] profiloPQNR = nomiDefiniti[DefinedNames.GetName(up, "PQNR_PROFILO", suffissoData)];
                    object[,] values = ws.Range[ws.Cells[profiloPQNR[0].Item1, profiloPQNR[0].Item2], ws.Cells[profiloPQNR[0].Item1, profiloPQNR[profiloPQNR.Length - 1].Item2]].Value;
                    object[] valoriPQNR = values.Cast<object>().ToArray();

                    DataView assetti = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAASSETTO].DefaultView;
                    assetti.RowFilter = "SiglaEntita = '" + up + "'";

                    double?[] pMin = new double?[valoriPQNR.Length];
                    int numAssetto = 1;
                    foreach (DataRowView assetto in assetti)
                    {
                        Tuple<int, int>[] cellePmin = nomiDefiniti[DefinedNames.GetName(up, "PMIN_TERNA_ASSETTO" + numAssetto, suffissoData)];
                        object[,] pMinAssetto = ws.Range[ws.Cells[cellePmin[0].Item1, cellePmin[0].Item2], ws.Cells[cellePmin[0].Item1, cellePmin[cellePmin.Length - 1].Item2]].Value;
                        double?[] pMinOraria = pMinAssetto.Cast<double?>().ToArray();
                        for (int i = 0; i < pMinOraria.Length; i++)
                        {
                            pMin[i] = Math.Min(pMin[i] ?? pMinOraria[i] ?? 0, pMinOraria[i] ?? 0);
                        }
                        numAssetto++;
                    }

                    int oreGiorno = valoriPQNR.Length;

                    int oreFermata = int.Parse(CommonFunctions.DB.Select("spGetOreFermata", "@SiglaEntita=" + up).Rows[0]["OreFermata"].ToString());

                    Forms.frmRAMPE rampe = new frmRAMPE(desEntita, pRif, pMin, oreGiorno, entitaRampa, valoriPQNR, oreFermata);
                    rampe.Text = Simboli.nomeApplicazione;
                    rampe.ShowDialog();

                    if (rampe._out != null)
                    {
                        ws.Range[ws.Cells[profiloPQNR[0].Item1, profiloPQNR[0].Item2], ws.Cells[profiloPQNR[0].Item1, profiloPQNR[profiloPQNR.Length - 1].Item2]].Value = rampe._out.AsEnumerable().Select(r => r["SiglaRampa"]).ToArray();

                        for (int i = 1; i < rampe._out.Columns.Count; i++)
                        {
                            Tuple<int, int>[] pqnrX = nomiDefiniti[DefinedNames.GetName(up, "PQNR" + i, suffissoData)];
                            ws.Range[ws.Cells[pqnrX[0].Item1, pqnrX[0].Item2], ws.Cells[pqnrX[0].Item1, pqnrX[pqnrX.Length - 1].Item2]].Value = rampe._out.AsEnumerable().Select(r => r["Q" + i]).ToArray();
                        }
                    }
                }
            }
            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }
        private void btnAggiornaDati_Click(object sender, RibbonControlEventArgs e)
        {            
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            if (CommonFunctions.DB.OpenConnection() && CommonFunctions.DB.StatoDB()[DataBase.NomiDB.SQLSERVER] == ConnectionState.Open)
            {
                AggiornaDati();

                //TODO riabilitare log!!
                //CommonFunctions.InsertLog(DataBase.TipologiaLOG.LogModifica, "Aggiorna Dati");

                CommonFunctions.DB.CloseConnection();
            }
            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }
        private void btnAzioni_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Sheet.Proteggi(false);
            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;

            frmAZIONI frmAz = new frmAZIONI();
            frmAz.Text = Simboli.nomeApplicazione;
            frmAz.ShowDialog();

            Globals.ThisWorkbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Sheet.Proteggi(true);
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }
        private void btnModifica_Click(object sender, RibbonControlEventArgs e)
        {
            Simboli.ModificaDati = btnModifica.Checked;

            Sheet.AbilitaModifica(btnModifica.Checked);
            if (btnModifica.Checked)
                btnModifica.Label = "Modifica SI";
            else
                btnModifica.Label = "Modifica NO";
        }

        #endregion

        #region Metodi

        private void AbilitaTasti(bool abilita)
        {
            btnCalendar.Enabled = abilita;
            btnAzioni.Enabled = abilita;
            btnRampe.Enabled = abilita;
            btnAggiornaDati.Enabled = abilita;
        }
        
        private void AggiornaStruttura()
        {
            CommonFunctions.AggiornaStrutturaDati();

            DataView categorie = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIA].DefaultView;
            categorie.RowFilter = "Operativa = 1";

            foreach (DataRowView categoria in categorie)
            {
                Excel.Worksheet ws;
                try
                {
                    ws = Globals.ThisWorkbook.Worksheets[categoria["DesCategoria"].ToString()];
                }
                catch
                {
                    ws = (Excel.Worksheet)Globals.ThisWorkbook.Worksheets.Add(Globals.ThisWorkbook.Worksheets["Log"]);
                    ws.Name = categoria["DesCategoria"].ToString();
                    ws.Select();
                    Globals.ThisWorkbook.Application.Windows[1].DisplayGridlines = false;                    
                }
            }

            Riepilogo main = new Riepilogo(Globals.ThisWorkbook.Sheets["Main"]);
            main.LoadStructure();

            foreach (Excel.Worksheet ws in Globals.ThisWorkbook.Sheets)
            {
                if (ws.Name != "Log" && ws.Name != "Main")
                {
                    Sheet s = new Sheet(ws);
                    s.LoadStructure();
                }
            }

            Globals.Main.Select();
            Globals.ThisWorkbook.Application.WindowState = Excel.XlWindowState.xlMaximized;
        }
        private void AggiornaDati()
        {
            foreach (Excel.Worksheet ws in Globals.ThisWorkbook.Sheets)
            {
                if (ws.Name != "Log" && ws.Name != "Main")
                {
                    Sheet s = new Sheet(ws);
                    s.UpdateData();
                }
            }
            Riepilogo main = new Riepilogo(Globals.ThisWorkbook.Sheets["Main"]);
            main.LoadStructure();

            //Log
            CommonFunctions.InitLog();
        }

        #endregion

    }
}
