using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    public abstract class ACarica
    {
        public abstract bool AzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime giorno, object parametro = null);
        /*public abstract void AzzeraInformazione(object siglaEntita, object siglaAzione, DefinedNames nomiDefiniti, DateTime giorno);
        public abstract void ScriviInformazione(object siglaEntita, DataView azioneInformazione, DefinedNames nomiDefiniti);*/
    }

    public class Carica : ACarica
    {
        public override bool AzioneInformazione(object siglaEntita, object siglaAzione, object azionePadre, DateTime giorno, object parametro = null)
        {
            DefinedNames definedNames = new DefinedNames(DefinedNames.GetSheetName(siglaEntita));
            try
            {
                AzzeraInformazione(siglaEntita, siglaAzione, definedNames, giorno);

                if (DataBase.OpenConnection())
                {
                    if (azionePadre.Equals("GENERA"))
                    {
                        ElaborazioneInformazione(siglaEntita, siglaAzione, definedNames, giorno);
                        DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno);
                    }
                    else
                    {
                        DataView azioneInformazione = DataBase.Select(DataBase.SP.CARICA_AZIONE_INFORMAZIONE, "@SiglaEntita=" + siglaEntita + ";@SiglaAzione=" + siglaAzione + ";@Parametro=" + parametro + ";@Data=" + giorno.ToString("yyyyMMdd")).DefaultView;
                        if (azioneInformazione.Count == 0)
                        {
                            DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno, false);
                            return false;
                        }
                        else
                        {
                            ScriviInformazione(siglaEntita, azioneInformazione, definedNames);
                            DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, giorno);
                        }
                    }

                    Sheet s = new Sheet(Workbook.Sheets[definedNames.Sheet]);
                    s.AggiornaGrafici();
                    return true;
                }
                else
                {
                    if (azionePadre.Equals("GENERA"))
                        ElaborazioneInformazione(siglaEntita, siglaAzione, definedNames, giorno);
                    else if (azionePadre.Equals("CARICA")) 
                    { /*TODO per invio programmi un caricamento da XML*/ }

                    Sheet s = new Sheet(Workbook.Sheets[definedNames.Sheet]);
                    s.AggiornaGrafici();
                    return true;
                }
            }
            catch (Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "CaricaAzioneInformazione [" + siglaEntita + ", " + siglaAzione + "]: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
        
        protected virtual void AzzeraInformazione(object siglaEntita, object siglaAzione, DefinedNames definedNames, DateTime giorno)
        {
            Excel.Worksheet ws = Workbook.Sheets[definedNames.Sheet];

            string suffissoData = Date.GetSuffissoData(giorno);

            DataView azioneInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
            azioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

            foreach (DataRowView info in azioneInformazione)
            {
                if (info["FormulaInCella"].Equals("0"))
                {
                    siglaEntita = info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"];
                    Range rng;
                    //if(info["Selezione"].Equals(0))
                    rng = definedNames.Get(siglaEntita, info["SiglaInformazione"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));
                    //else
                    //    rng = nomiDefiniti.Get(siglaEntita, "SEL", info["SiglaInformazione"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));

                    Excel.Range xlRng = ws.Range[rng.ToString()];
                    xlRng.Value = null;
                    Style.RangeStyle(xlRng, backColor: info["BackColor"], foreColor: info["ForeColor"]);
                    xlRng.ClearComments();
                }
            }
        }
        protected virtual void ScriviInformazione(object siglaEntita, DataView azioneInformazione, DefinedNames definedNames)
        {
            Excel.Worksheet ws = Workbook.Sheets[definedNames.Sheet];

            foreach (DataRowView azione in azioneInformazione)
            {
                string suffissoData;
                string suffissoOra;
                if (azione["SiglaEntita"].Equals("UP_BUS") && azione["SiglaInformazione"].Equals("VOL_INVASO"))
                {
                    suffissoData = Date.GetSuffissoData(DataBase.DataAttiva.AddDays(-1));
                    suffissoOra = Date.GetSuffissoOra(24);
                }
                else
                {
                    suffissoData = Date.GetSuffissoData(DataBase.DataAttiva, azione["Data"]);
                    suffissoOra = Date.GetSuffissoOra(azione["Data"]);
                }

                ScriviCella(ws, definedNames, azione["SiglaEntita"], azione, suffissoData, suffissoOra, azione["Valore"], false);
            }
        }
        protected void ElaborazioneInformazione(object siglaEntita, object siglaAzione, DefinedNames definedNames, DateTime giorno, int oraInizio = -1, int oraFine = -1)
        {
            Excel.Worksheet ws = Workbook.Sheets[definedNames.Sheet];

            Dictionary<string, int> entitaRiferimento = new Dictionary<string, int>();
            List<int> oreDaCalcolare = new List<int>();

            string suffissoData = Date.GetSuffissoData(giorno);

            oraInizio = oraInizio < 0 ? 1 : oraInizio;
            oraFine = oraFine < 0 ? Date.GetOreGiorno(giorno) : oraFine;

            DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "Gerarchia = '" + siglaEntita + "'";
            foreach (DataRowView entita in categoriaEntita)
                entitaRiferimento.Add(entita["SiglaEntita"].ToString(), (int)entita["Riferimento"]);

            if (entitaRiferimento.Count == 0)
                entitaRiferimento.Add(siglaEntita.ToString(), 1);

            DataView calcoloInformazione = DataBase.LocalDB.Tables[DataBase.Tab.CALCOLO_INFORMAZIONE].DefaultView;

            DataView entitaAzioneCalcolo = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE_CALCOLO].DefaultView;
            entitaAzioneCalcolo.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
            foreach (DataRowView azioneCalcolo in entitaAzioneCalcolo)
            {
                calcoloInformazione.RowFilter = "SiglaCalcolo = '" + azioneCalcolo["SiglaCalcolo"] + "'";
                calcoloInformazione.Sort = "Step";

                //azzero tutte le informazioni che vengono utilizzate nel calcolo tranne i CHECK
                foreach (DataRowView info in calcoloInformazione)
                {
                    if (!info["SiglaInformazione"].Equals("CHECKINFO"))
                    {
                        Range rng = definedNames.Get(info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"], info["SiglaInformazione"], suffissoData, Date.GetSuffissoOra(oraInizio)).Extend(colOffset: oraFine - oraInizio + 1);
                        ws.Range[rng.ToString()].Value = null;
                    }
                }

                for (int ora = oraInizio; ora <= oraFine; ora++)
                {
                    int i = 0;
                    while (i < calcoloInformazione.Count)
                    {
                        DataRowView calcolo = calcoloInformazione[i];
                        if (calcolo["OraInizio"] != DBNull.Value)
                            if (ora < int.Parse(calcolo["OraInizio"].ToString()) || ora > int.Parse(calcolo["OraFine"].ToString()))
                            {
                                i++;
                                continue;
                            }

                        if (calcolo["OraFine"] != DBNull.Value)
                            if (ora != Date.GetOreGiorno(giorno))
                                if (calcolo["FineCalcolo"].Equals("1"))
                                {
                                    i++;
                                    continue;
                                }
                                else
                                    break;

                        int step = 0;
                        object risultato = GetRisultatoCalcolo(siglaEntita, definedNames, giorno, ora, calcolo, entitaRiferimento, out step);

                        if (step == 0)
                        {
                            ScriviCella(ws, definedNames, siglaEntita, calcolo, suffissoData, Date.GetSuffissoOra(ora), risultato, true);
                        }

                        if (calcolo["FineCalcolo"].Equals("1") || step == -1)
                            break;

                        if (calcolo["GoStep"] != DBNull.Value)
                            step = (int)calcolo["GoStep"];

                        if (step != 0)
                            i = calcoloInformazione.Find(step);
                        else
                            i++;
                    }
                }
            }
        }
        protected virtual void ScriviCella(Excel.Worksheet ws, DefinedNames definedNames, object siglaEntita, DataRowView info, string suffissoData, string suffissoOra, object risultato, bool saveToDB) 
        {
            object siglaEntitaRif = siglaEntita;

            if(info.DataView.Table.Columns.Contains("SiglaEntitaRif") && info["SiglaEntitaRif"] != DBNull.Value)
                info["SiglaEntitaRif"] = info["SiglaEntitaRif"];

            Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], suffissoData, suffissoOra);
            Excel.Range xlRng = ws.Range[rng.ToString()];

            xlRng.Value = risultato;

            if (info["BackColor"] != DBNull.Value)
                xlRng.Interior.ColorIndex = info["BackColor"];
            if (info["ForeColor"] != DBNull.Value)
                xlRng.Font.ColorIndex = info["ForeColor"];

            xlRng.ClearComments();

            if (info["Commento"] != DBNull.Value)
                xlRng.AddComment(info["Commento"]).Visible = false;

            if(saveToDB)
                Handler.StoreEdit(xlRng, 0, true);
        }

        protected object GetRisultatoCalcolo(object siglaEntita, DefinedNames definedNames, DateTime giorno, int ora, DataRowView calcolo, Dictionary<string, int> entitaRiferimento, out int step)
        {
            Excel.Worksheet ws = Workbook.Sheets[definedNames.Sheet];

            string suffissoData = Date.GetSuffissoData(giorno);

            int ora1 = calcolo["OraInformazione1"] is DBNull ? ora : ora + (int)calcolo["OraInformazione1"];
            int ora2 = calcolo["OraInformazione2"] is DBNull ? ora : ora + (int)calcolo["OraInformazione2"];

            object siglaEntitaRif1 = calcolo["Riferimento1"] is DBNull ? (calcolo["SiglaEntita1"] is DBNull ? siglaEntita : calcolo["SiglaEntita1"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == (int)calcolo["Riferimento1"]).Key;
            object siglaEntitaRif2 = calcolo["Riferimento2"] is DBNull ? (calcolo["SiglaEntita2"] is DBNull ? siglaEntita : calcolo["SiglaEntita2"]) : entitaRiferimento.FirstOrDefault(kv => kv.Value == (int)calcolo["Riferimento2"]).Key;

            object valore1 = 0d;
            object valore2 = 0d;

            if (calcolo["SiglaInformazione1"] != DBNull.Value)
            {
                try
                {
                    Range cella1 = definedNames.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1));

                    switch (calcolo["SiglaInformazione1"].ToString())
                    {
                        case "UNIT_COMM":
                            DataView entitaCommitment = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_COMMITMENT].DefaultView;
                            entitaCommitment.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' AND SiglaCommitment = '" + ws.Range[cella1.ToString()].Value + "'";
                            valore1 = entitaCommitment.Count > 0 ? entitaCommitment[0]["IdEntitaCommitment"] : null;

                            break;
                        case "DISPONIBILITA":
                            if (ws.Range[cella1.ToString()].Value == "OFF")
                                valore1 = 0d;
                            else
                                valore1 = 1d;

                            break;
                        case "CHECKINFO":
                            if (ws.Range[cella1.ToString()].Value == "OK")
                                valore1 = 1d;
                            else
                                valore1 = 2d;
                            break;
                        default:
                            //if (cella != null)
                            valore1 = ws.Range[cella1.ToString()].Value ?? 0d;
                            break;
                    }
                }
                catch
                {
                    valore1 = 0d;
                }
            }
            else if (calcolo["IdProprieta"] != DBNull.Value)
            {
                DataView entitaProprieta = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' AND IdProprieta = " + calcolo["IdProprieta"];

                if (entitaProprieta.Count > 0)
                    valore1 = entitaProprieta[0]["Valore"];
            }
            else if (calcolo["IdParametroD"] != DBNull.Value)
            {
                DataView entitaParametro = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_D].DefaultView;
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' AND IdParametro = " + calcolo["IdParametroD"];

                if (entitaParametro.Count > 0)
                    valore1 =entitaParametro[0]["Valore"];
            }
            else if (calcolo["IdParametroH"] != DBNull.Value)
            {
                DataView entitaParametro = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PARAMETRO_H].DefaultView;
                entitaParametro.RowFilter = "SiglaEntita = '" + siglaEntitaRif1 + "' AND IdParametro = " + calcolo["IdParametroH"];

                if (entitaParametro.Count > 0)
                    valore1 = entitaParametro[0]["Valore"];
            }
            else if (calcolo["Valore"] != DBNull.Value)
            {
                valore1 = calcolo["Valore"];
            }

            if (calcolo["SiglaInformazione2"] != DBNull.Value)
            {
                try
                {
                    Range cella2 = definedNames.Get(siglaEntitaRif2, calcolo["SiglaInformazione2"], suffissoData, Date.GetSuffissoOra(ora2));

                    switch (calcolo["SiglaInformazione2"].ToString())
                    {
                        case "UNIT_COMM":
                            DataView entitaCommitment = DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_COMMITMENT].DefaultView;
                            entitaCommitment.RowFilter = "SiglaEntita = '" + siglaEntitaRif2 + "' AND SiglaCommitment = '" + ws.Range[cella2.ToString()].Value + "'";
                            valore2 = entitaCommitment.Count > 0 ? entitaCommitment[0] : null;

                            break;
                        case "DISPONIBILITA":
                            if (ws.Range[cella2.ToString()].Value == "OFF")
                                valore2 = 0d;
                            else
                                valore2 = 1d;

                            break;
                        case "CHECKINFO":
                            if (ws.Range[cella2.ToString()].Value == "OK")
                                valore2 = 1d;
                            else
                                valore2 = 2d;
                            break;
                        default:
                            //if (cella != null)
                            valore2 = ws.Range[cella2.ToString()].Value ?? 0d;
                            //else
                            //    valore2 = 0d;
                            break;
                    }
                }
                catch
                {
                    valore2 = 0d;
                }

            }

            double retVal = 0d;

            valore1 = valore1 ?? 0d;
            valore2 = valore2 ?? 0d;

            if (calcolo["Funzione"] is DBNull && calcolo["Operazione"] is DBNull && calcolo["Condizione"] is DBNull)
            {
                step = 0;
                if (Convert.ToDouble(valore1) == 0d)
                    return valore2;

                return valore1;
            }
            else if (calcolo["Funzione"] != DBNull.Value)
            {
                string func = calcolo["Funzione"].ToString().ToLowerInvariant();
                if (calcolo["SiglaInformazione2"] is DBNull)
                {
                    if (func.Contains("abs"))
                    {
                        retVal = Math.Abs(Convert.ToDouble(valore1));
                    }
                    else if (func.Contains("floor"))
                    {
                        retVal = Math.Floor(Convert.ToDouble(valore1));
                    }
                    else if (func.Contains("round"))
                    {
                        int decimals = int.Parse(func.Replace("round", ""));
                        retVal = Math.Round(Convert.ToDouble(valore1), decimals);
                    }
                    else if (func.Contains("power"))
                    {
                        int exp = int.Parse(Regex.Match(func, @"\d*").Value);
                        retVal = Math.Pow(Convert.ToDouble(valore1), exp);
                    }
                    else if (func.Contains("sum"))
                    {
                        foreach (var kvp in entitaRiferimento)
                            retVal += ws.Range[definedNames.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0d;
                    }
                    else if (func.Contains("avg"))
                    {
                        foreach (var kvp in entitaRiferimento)
                            retVal += ws.Range[definedNames.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0d;
                        retVal /= entitaRiferimento.Count;
                    }
                    else if (func.Contains("max_h"))
                    {
                        Range rng = definedNames.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));
                        object[,] tmpVal = ws.Range[rng.ToString()].Value;
                        for (int i = 1; i <= tmpVal.GetLength(1); i++)
                            if (tmpVal[1, i] == null)
                                tmpVal[1, i] = 0d;

                        double[] values = tmpVal.Cast<double>().ToArray();
                        retVal = values.Max();
                    }
                    else if (func.Contains("min_h"))
                    {
                        Range rng = definedNames.Get(siglaEntitaRif1, calcolo["SiglaInformazione1"], suffissoData).Extend(colOffset: Date.GetOreGiorno(giorno));
                        object[,] tmpVal = ws.Range[rng.ToString()].Value;
                        for (int i = 1; i <= tmpVal.GetLength(1); i++)
                            if (tmpVal[1, i] == null)
                                tmpVal[1, i] = 0d;

                        double[] values = tmpVal.Cast<double>().ToArray();
                        retVal = values.Min();
                    }
                    else if (func.Contains("max"))
                    {
                        retVal = double.MinValue;
                        foreach (var kvp in entitaRiferimento)
                            retVal = Math.Max(ws.Range[definedNames.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0, retVal);
                    }
                    else if (func.Contains("min"))
                    {
                        retVal = double.MaxValue;
                        foreach (var kvp in entitaRiferimento)
                            retVal = Math.Min(ws.Range[definedNames.Get(kvp.Key, calcolo["SiglaInformazione1"], suffissoData, Date.GetSuffissoOra(ora1)).ToString()].Value ?? 0, retVal);
                    }
                }
                //caso in cui ci sia anche SiglaInformazione2
                else
                {
                    if (func.Contains("max"))
                    {
                        retVal = Math.Max(Convert.ToDouble(valore1), Convert.ToDouble(valore2));
                    }
                    else if (func.Contains("min"))
                    {
                        retVal = Math.Min(Convert.ToDouble(valore1), Convert.ToDouble(valore2));
                    }
                }
            }
            else if (calcolo["Operazione"] != DBNull.Value)
            {
                switch (calcolo["Operazione"].ToString())
                {
                    case "+":
                        retVal = Convert.ToDouble(valore1) + Convert.ToDouble(valore2);
                        break;
                    case "-":
                        retVal = Convert.ToDouble(valore1) - Convert.ToDouble(valore2);
                        break;
                    case "*":
                        retVal = Convert.ToDouble(valore1) * Convert.ToDouble(valore2);
                        break;
                    case "/":
                        retVal = Convert.ToDouble(valore1) / Convert.ToDouble(valore2);
                        break;
                }
            }
            else if (calcolo["Condizione"] != DBNull.Value)
            {
                bool res = false;
                switch (calcolo["Condizione"].ToString())
                {
                    case ">":
                        res = Convert.ToDouble(valore1) > Convert.ToDouble(valore2);
                        break;
                    case "<":
                        res = Convert.ToDouble(valore1) < Convert.ToDouble(valore2);
                        break;
                    case ">=":
                        res = Convert.ToDouble(valore1) >= Convert.ToDouble(valore2);
                        break;
                    case "<=":
                        res = Convert.ToDouble(valore1) <= Convert.ToDouble(valore2);
                        break;
                    case "=":
                        res = Convert.ToDouble(valore1) == Convert.ToDouble(valore2);
                        break;
                    case "<>":
                        res = Convert.ToDouble(valore1) != Convert.ToDouble(valore2);
                        break;
                }
                if (res)
                    step = (int)calcolo["StepCondizioneVera"];
                else
                    step = (int)calcolo["StepCondizioneFalsa"];

                return res;
            }

            step = 0;
            return retVal;
        }
    }
}
