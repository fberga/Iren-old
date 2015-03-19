using Iren.ToolsExcel.UserConfig;
using Iren.ToolsExcel.Utility;
using Iren.ToolsExcel.Base;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    public class Esporta : Base.Esporta
    {
        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            DataView entitaAzione = _localDB.Tables[Utility.DataBase.Tab.ENTITAAZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
            if (entitaAzione.Count == 0)
                return false;

            DataView categoriaEntita = _localDB.Tables[Utility.DataBase.Tab.CATEGORIAENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
            object codiceRUP = categoriaEntita[0]["CodiceRUP"];

            DataView entitaAzioneInformazione = _localDB.Tables[Utility.DataBase.Tab.ENTITAAZIONEINFORMAZIONE].DefaultView;
            entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
            Excel.Worksheet ws = Workbook.WB.Sheets[nomeFoglio];

            switch (siglaAzione.ToString())
            {
                case "E_PREZZO_MSD":
                    DataTable dt = new DataTable("E_PREZZO_MSD")
                    {
                        Columns =
                        {
                            {"UP", typeof(string)},
                            {"Data", typeof(string)},
                            {"Ora", typeof(string)},
                            {"PREZZO_MSD_MINIMO", typeof(string)},
                            {"PREZZO_MSD_SPEGNIMENTO", typeof(string)},
                            {"PREZZO_MSD_AS1_V", typeof(string)},
                            {"PREZZO_MSD_AS1_A", typeof(string)},
                            {"PREZZO_MSD_AS2_V", typeof(string)},
                            {"PREZZO_MSD_AS2_A", typeof(string)},
                            {"PREZZO_MSD_AS3_V", typeof(string)},
                            {"PREZZO_MSD_AS3_A", typeof(string)},
                            {"PREZZO_MSD_RS_V", typeof(string)},
                            {"PREZZO_MSD_RS_A", typeof(string)},
                            {"ACCENSIONE", typeof(string)},
                            {"CAMBIO_ASSETTO", typeof(string)}
                        }
                    };

                    bool valAccensioneNULL = false;
                    bool cambioAssettoNULL = false;
                    bool datiEsportatiAccensioneNULL = false;
                    bool datiEsportatiCambioAssettoNULL = false;

                    bool isDefinedAccensione = nomiDefiniti.IsDefined(DefinedNames.GetName(siglaEntita, "ACCENSIONE", "DATA1"));
                    bool isDefinedCambioAssetto = nomiDefiniti.IsDefined(DefinedNames.GetName(siglaEntita, "CAMBIO_ASSETTO", "DATA1"));

                    for (DateTime giorno = DataBase.DataAttiva, dataFine = DataBase.DataAttiva.AddDays(Struct.intervalloGiorni); giorno <= dataFine; giorno = giorno.AddDays(1))
                    {
                        string suffissoData = Date.GetSuffissoData(DataBase.DataAttiva, giorno);
                        int oreData = Date.GetOreGiorno(giorno);

                        object valAccensione = "NULL";
                        if (isDefinedAccensione)
                        {
                            Tuple<int,int> cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "ACCENSIONE", suffissoData)][0];
                            valAccensione = ws.Cells[cella.Item1, cella.Item2].Value;
                            if (valAccensione == null)
                            {
                                valAccensioneNULL = true;
                                valAccensione = "0";
                            }
                        }

                        object cambioAssetto = "NULL";
                        if (isDefinedCambioAssetto)
                        {
                            Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "CAMBIO_ASSETTO", suffissoData)][0];
                            cambioAssetto = ws.Cells[cella.Item1, cella.Item2].Value;
                            if (cambioAssetto == null)
                            {
                                cambioAssettoNULL = true;
                                cambioAssetto = "0";
                            } 
                        }

                        Dictionary<object, object[]> informazioni = new Dictionary<object, object[]>();
                        foreach (DataRowView entAzInfo in entitaAzioneInformazione)
                        {
                            if (nomiDefiniti.IsDefined(DefinedNames.GetName(siglaEntita, entAzInfo["SiglaInformazione"], suffissoData)))
                            {
                                Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(siglaEntita, entAzInfo["SiglaInformazione"], suffissoData)];
                                object[,] tmp = ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]].Value;
                                object[] tmp1 = tmp.Cast<object>().ToArray();
                                object notNull = Array.Find<object>(tmp1, obj => obj != null);
                                if(notNull != null)
                                    informazioni.Add(entAzInfo["SiglaInformazione"], tmp1);
                            }
                        }

                        if (informazioni.Count > 0)
                        {
                            if (isDefinedAccensione && valAccensioneNULL)
                                datiEsportatiAccensioneNULL = true;

                            if (isDefinedCambioAssetto && cambioAssettoNULL)
                                datiEsportatiCambioAssettoNULL = true;

                            for (int i = 0; i < oreData; i++)
                            {
                                DataRow row = dt.NewRow();

                                row["UP"] = codiceRUP;
                                row["Data"] = giorno.ToString("dd/MM/yyyy");
                                row["Ora"] = (i + 1).ToString("00");

                                int startIndex = dt.Columns.IndexOf("PREZZO_MSD_MINIMO");
                                int maxIndex = dt.Columns.IndexOf("ACCENSIONE");
                                for (int j = startIndex; j < maxIndex; j++)
                                {
                                    if (informazioni.ContainsKey(dt.Columns[j].ColumnName))
                                    {
                                        object val = informazioni[dt.Columns[j].ColumnName][i];
                                        row[j] = val == null ? "NULL" : val;
                                    }
                                    else
                                        row[j] = "NULL";
                                }
                                row["ACCENSIONE"] = valAccensione.ToString();
                                row["CAMBIO_ASSETTO"] = cambioAssetto.ToString();

                                dt.Rows.Add(row);
                            }
                        }
                    }

                    if (datiEsportatiAccensioneNULL)
                        System.Windows.Forms.MessageBox.Show("Per " + codiceRUP + " sono stati esportati valori di Accensione nulli", Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                    if (datiEsportatiCambioAssettoNULL)
                        System.Windows.Forms.MessageBox.Show("Per " + codiceRUP + " sono stati esportati valori di Cambio Assetto nulli", Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);

                    var path = Utility.Utilities.GetUsrConfigElement("pathCaricatorePEXCA");

                    string pathStr = Utility.ExportPath.PreparePath(path.Value);

                    if (Directory.Exists(pathStr))
                    {
                        if (!ExportToCSV(System.IO.Path.Combine(pathStr, "PREZZO_MSD_" + codiceRUP + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfffffff") + ".csv"), dt))
                            return false;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Il percorso '" + pathStr + "' non è raggiungibile.", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return false;
                    }

                    break;
            }
            return true;
        }
    }
}
