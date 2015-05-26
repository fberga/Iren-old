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
            DataView entitaAzione = _localDB.Tables[Utility.DataBase.Tab.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
            if (entitaAzione.Count == 0)
                return false;

            DataView categoriaEntita = _localDB.Tables[Utility.DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
            object codiceRUP = categoriaEntita[0]["CodiceRUP"];

            DataView entitaProprieta = _localDB.Tables[Utility.DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'IMP_COD_IF'";
            object codiceIF = entitaProprieta[0]["Valore"];

            DataView entitaAzioneInformazione = _localDB.Tables[Utility.DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
            entitaAzioneInformazione.RowFilter = "SiglaEntitaRif = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames definedNames = new DefinedNames(nomeFoglio, DefinedNames.InitType.NamingOnly);

            Excel.Worksheet ws = Workbook.WB.Sheets[nomeFoglio];

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

                    string suffissoData = Utility.Date.GetSuffissoData(dataRif);
                    int oreData = Utility.Date.GetOreGiorno(dataRif);
                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        object siglaEntitaRif = (info["SiglaEntitaRif"] is DBNull ? siglaEntita: info["SiglaEntitaRif"]);

                        Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], suffissoData).Extend(colOffset: oreData);

                        //object[,] values = ws.Range[rng.ToString()].Value;
                        //bool empty = true;
                        //foreach (object value in values)
                        //{
                        //    if(value != null) 
                        //    {
                        //        empty = false;
                        //        break;
                        //    }
                        //}

                        //if (!empty)
                        //{
                            for (int i = 0; i < rng.Columns.Count; i++)
                            {
                                DataRow row = dt.NewRow();

                                row["Campo1"] = "ASSET";
                                row["Campo2"] = "Produzione";
                                row["UP"] = codiceIF;
                                row["Campo3"] = "NA";
                                row["Data"] = dataRif.ToString("dd/MM/yyyy");
                                row["Ora"] = (i + 1).ToString("00") + ".00";
                                row["Campo4"] = "ASSETTO";
                                row["UnitComm"] = ws.Range[rng.Columns[i].ToString()].Value;
                                row["Campo5"] = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

                                dt.Rows.Add(row);
                            }
                        //}
                    }

                    var path = Utility.Workbook.GetUsrConfigElement("pathCaricatoreImpianti");

                    string pathStr = Utility.ExportPath.PreparePath(path.Value);

                    if (Directory.Exists(pathStr))
                    {
                        if (dt.AsEnumerable().Any(r => r["UnitComm"] != DBNull.Value)
                            && ExportToCSV(System.IO.Path.Combine(pathStr, "AEM_ASSET_" + codiceIF + "_" + dataRif.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfffffff") + ".csv"), dt))
                            return true;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Il percorso '" + pathStr + "' non è raggiungibile.", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    }
                    break;
            }
            return false;
        }
    }
}
