using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    public class Carica : Base.Carica
    {
        DefinedNames _definedNamesSheetMercato = new DefinedNames("MSD1");  //non mi interessa sapere il mercato... sono tutti uguali
        Excel.Worksheet _wsMercato;

        public Carica() 
            : base() 
        {
            _wsMercato = Workbook.Sheets[Simboli.Mercato];
        }

        protected override void ScriviInformazione(object siglaEntita, DataView azioneInformazione, DefinedNames definedNames)
        {
            Excel.Worksheet ws = Workbook.Sheets[definedNames.Sheet];

            DataTable entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA];

            var rif =
                (from r in entita.AsEnumerable()
                 where r["SiglaEntita"].Equals(siglaEntita)
                 select new { SiglaEntita = r["Gerarchia"] is DBNull ? r["SiglaEntita"] : r["Gerarchia"], Riferimento = r["Riferimento"] }).First();

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

                Range rng = definedNames.Get(siglaEntita, azione["SiglaInformazione"], suffissoData, suffissoOra);

                Excel.Range xlRng = ws.Range[rng.ToString()];
                xlRng.Value = azione["Valore"];
                if (azione["BackColor"] != DBNull.Value)
                    xlRng.Interior.ColorIndex = azione["BackColor"];
                if (azione["BackColor"] != DBNull.Value)
                    xlRng.Font.ColorIndex = azione["ForeColor"];

                xlRng.ClearComments();

                if (azione["Commento"] != DBNull.Value)
                    xlRng.AddComment(azione["Commento"]).Visible = false;


                //copio le informazioni sui fogli dei mercati

                string quarter = Regex.Match(azione["SiglaInformazione"].ToString(), @"Q\d").Value;
                quarter = quarter == "" ? "Q1" : quarter;

                Range rngMercato = new Range(_definedNamesSheetMercato.GetRowByName(rif.SiglaEntita, "UM", "T") + 2, _definedNamesSheetMercato.GetColFromName("RIF" + rif.Riferimento, "PROGRAMMA" + quarter));
                rngMercato.StartRow += (Date.GetOraFromDataOra(azione["Data"].ToString()) - 1);

                _wsMercato.Range[rngMercato.ToString()].Value = xlRng.Value;
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
                                continue;

                        if (calcolo["OraFine"] != DBNull.Value)
                            if (ora != Date.GetOreGiorno(giorno))
                                if (calcolo["FineCalcolo"].Equals("1"))
                                    continue;
                                else
                                    break;

                        int step = 0;
                        object risultato = GetRisultatoCalcolo(siglaEntita, definedNames, giorno, ora, calcolo, entitaRiferimento, out step);

                        if (step == 0)
                        {
                            object siglaEntitaRif = calcolo["SiglaEntitaRif"] is DBNull ? siglaEntita : calcolo["SiglaEntitaRif"];
                            Range rng = definedNames.Get(siglaEntitaRif, calcolo["SiglaInformazione"], suffissoData, Date.GetSuffissoOra(ora));
                            Excel.Range xlRng = ws.Range[rng.ToString()];

                            xlRng.Value = risultato;

                            if (calcolo["BackColor"] != DBNull.Value)
                                xlRng.Interior.ColorIndex = calcolo["BackColor"];
                            if (calcolo["ForeColor"] != DBNull.Value)
                                xlRng.Font.ColorIndex = calcolo["ForeColor"];

                            xlRng.ClearComments();

                            if (calcolo["Commento"] != DBNull.Value)
                                xlRng.AddComment(calcolo["Commento"]).Visible = false;

                            Handler.StoreEdit(xlRng, 0);
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
    }
}
