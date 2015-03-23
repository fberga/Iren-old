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

            DataView entitaProprieta = _localDB.Tables[Utility.DataBase.Tab.ENTITAPROPRIETA].DefaultView;
            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'IMP_COD_IF'";
            object codiceIF = entitaProprieta[0]["Valore"];

            DataView entitaAzioneInformazione = _localDB.Tables[Utility.DataBase.Tab.ENTITAAZIONEINFORMAZIONE].DefaultView;
            entitaAzioneInformazione.RowFilter = "SiglaEntitaRif = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

            switch (siglaAzione.ToString())
            {
                case "E_UNIT_COMM":
                    DataTable dt = new DataTable("E_UNIT_COMM")
                    {
                        Columns =
                        {
                            {"Campo1", typeof(string)},
                            {"Campo2", typeof(string)},
                            {"UP", typeof(string)},
                            {"Campo3", typeof(string)},
                            {"Data", typeof(string)},
                            {"Ora", typeof(string)},
                            {"Campo4", typeof(string)},
                            {"UnitComm", typeof(string)},
                            {"Campo5", typeof(string)}
                        }
                    };

                    string suffissoData = Utility.Date.GetSuffissoData(_db.DataAttiva, dataRif);
                    foreach (DataRowView entAzInfo in entitaAzioneInformazione)
                    {
                        object entita = (entAzInfo["SiglaEntitaRif"] is DBNull ? entAzInfo["SiglaEntita"] : entAzInfo["SiglaEntitaRif"]);

                        Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(entita, entAzInfo["SiglaInformazione"], suffissoData)];

                        string range = Sheet.R1C1toA1(riga[0].Item1, riga[0].Item2) + ":" + Sheet.R1C1toA1(riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2);

                        Excel.Worksheet ws = Workbook.WB.Sheets[nomeFoglio];
                        Excel.Range rng = ws.Range[range];
                        object[,] tmpVal = rng.Value;
                        object[] values = tmpVal.Cast<object>().ToArray();

                        for (int i = 0, length = values.Length; i < length; i++)
                        {
                            DataRow row = dt.NewRow();

                            row["Campo1"] = "ASSET";
                            row["Campo2"] = "Produzione";
                            row["UP"] = codiceIF;
                            row["Campo3"] = "NA";
                            row["Data"] = dataRif.ToString("dd/MM/yyyy");
                            row["Ora"] = (i + 1).ToString("00") + ".00";
                            row["Campo4"] = "ASSETTO";
                            row["UnitComm"] = values[i];
                            row["Campo5"] = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                            dt.Rows.Add(row);
                        }
                    }

                    var path = Utility.Utilities.GetUsrConfigElement("pathCaricatoreImpianti");

                    string pathStr = Utility.ExportPath.PreparePath(path.Value);

                    if (Directory.Exists(pathStr))
                    {
                        if (!ExportToCSV(System.IO.Path.Combine(pathStr, "AEM_ASSET_" + codiceIF + "_" + dataRif.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfffffff") + ".csv"), dt))
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
