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
        private void ToolsExcelRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            DateTime cfgDate = DateTime.ParseExact(ConfigurationManager.AppSettings["DataInizio"], "yyyyMMdd", CultureInfo.InvariantCulture);
            btnCalendar.Label = cfgDate.ToString("dddd dd MMM yyyy");
        }

        private void btnAggiornaStruttura_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
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
                if(ws.Name != "Log" && ws.Name != "Main")
                {
                    Sheet s = new Sheet(ws);
                    s.LoadStructure();
                }
            }

            Globals.Main.Select();
            Globals.ThisWorkbook.Application.WindowState = Excel.XlWindowState.xlMaximized;

            //TODO riabilitare log!!
            //Globals.Log.Unprotect();
            //CommonFunctions.InsertLog(DataBase.TipologiaLOG.LogModifica, "Aggiorna struttura");
            //Globals.Log.Protect();

            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        private void btnCalendar_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
            Forms.frmCALENDAR cal = new frmCALENDAR();
            cal.Text = Simboli.nomeApplicazione;
            cal.ShowDialog();

            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            if (cal.Date != null)
            {
                config.AppSettings.Settings["DataInizio"].Value = cal.Date.Value.ToString("yyyyMMdd");
                config.Save(ConfigurationSaveMode.Minimal);
                ConfigurationManager.RefreshSection("appSettings");

                btnCalendar.Label = cal.Date.Value.ToString("dddd dd MMM yyyy");
            }
            cal.Dispose();
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }

        private void btnRampe_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisWorkbook.Application.ScreenUpdating = false;
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

                    Tuple<int, int>[] profiloPQNR = nomiDefiniti[CommonFunctions.GetName(up, "PQNR_PROFILO", suffissoData)];
                    object[,] values = ws.Range[ws.Cells[profiloPQNR[0].Item1, profiloPQNR[0].Item2], ws.Cells[profiloPQNR[0].Item1, profiloPQNR[profiloPQNR.Length - 1].Item2]].Value;
                    object[] valoriPQNR = values.Cast<object>().ToArray();

                    DataView assetti = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAASSETTO].DefaultView;
                    assetti.RowFilter = "SiglaEntita = '" + up + "'";
                    
                    double?[] pMin = new double?[valoriPQNR.Length];
                    int numAssetto = 1;
                    foreach (DataRowView assetto in assetti)
                    {
                        Tuple<int,int>[] cellePmin = nomiDefiniti[CommonFunctions.GetName(up, "PMIN_TERNA_ASSETTO" + numAssetto, suffissoData)];
                        object[,] pMinAssetto = ws.Range[ws.Cells[cellePmin[0].Item1, cellePmin[0].Item2], ws.Cells[cellePmin[0].Item1, cellePmin[cellePmin.Length - 1].Item2]].Value;
                        double[] pMinOraria = pMinAssetto.Cast<double>().ToArray();
                        for (int i = 0; i < pMinOraria.Length; i++)
                        {
                            pMin[i] = Math.Min(pMin[i] ?? pMinOraria[i], pMinOraria[i]);
                        }
                        numAssetto++;
                    }

                    int oreGiorno = valoriPQNR.Length;

                    int oreFermata = int.Parse(CommonFunctions.DB.Select("spGetOreFermata", "@SiglaEntita=" + up).Rows[0]["OreFermata"].ToString());

                    Forms.frmRAMPE rampe = new frmRAMPE(desEntita, pRif, pMin, oreGiorno, entitaRampa, valoriPQNR, oreFermata);
                    rampe.ShowDialog();

                    if (rampe._out != null)
                    {
                        ws.Range[ws.Cells[profiloPQNR[0].Item1, profiloPQNR[0].Item2], ws.Cells[profiloPQNR[0].Item1, profiloPQNR[profiloPQNR.Length - 1].Item2]].Value = rampe._out.AsEnumerable().Select(r => r["SiglaRampa"]).ToArray();

                        for (int i = 1; i < rampe._out.Columns.Count; i++)
                        {
                            Tuple<int, int>[] pqnrX = nomiDefiniti[CommonFunctions.GetName(up, "PQNR" + i, suffissoData)];
                            ws.Range[ws.Cells[pqnrX[0].Item1, pqnrX[0].Item2], ws.Cells[pqnrX[0].Item1, pqnrX[pqnrX.Length - 1].Item2]].Value = rampe._out.AsEnumerable().Select(r => r["Q" + i]).ToArray();
                        }                         
                    }
                }
            }
            Globals.ThisWorkbook.Application.ScreenUpdating = true;
        }
    }
}
