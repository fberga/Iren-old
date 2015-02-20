﻿using Iren.FrontOffice.Core;
using Iren.FrontOffice.UserConfig;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.FrontOffice.Base
{
    public class Esporta : CommonFunctions, IEsporta
    {
        public Esporta()
        {
        }

        public bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime? dataRif = null)
        {
            if (dataRif == null)
                dataRif = DataBase.DataAttiva;

            try
            {
                DataView entitaAzione = LocalDB.Tables[Tab.ENTITAAZIONE].DefaultView;
                entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
                if (entitaAzione.Count == 0)
                    return false;

                DataView categoriaEntita = LocalDB.Tables[Tab.CATEGORIAENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];

                DataView entitaProprieta = LocalDB.Tables[Tab.ENTITAPROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'IMP_COD_IF'";
                object codiceIF = entitaProprieta[0]["Valore"];

                DataView entitaAzioneInformazione = LocalDB.Tables[Tab.ENTITAAZIONEINFORMAZIONE].DefaultView;
                entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

                switch (siglaAzione.ToString())
                {
                    case "E_MP_MGP":
                        DataTable dt = new DataTable("E_MP_MGP")
                        {
                            Columns =
                            {
                                {"Campo1", typeof(string)},
                                {"Campo2", typeof(string)},
                                {"UP", typeof(string)},
                                {"Campo3", typeof(string)},
                                {"Data", typeof(string)},
                                {"Ora", typeof(string)},
                                {"Informazione", typeof(string)},
                                {"Valore", typeof(string)}
                            }
                        };
                        
                        string suffissoData = GetSuffissoData(DataBase.DataAttiva, dataRif.Value);
                        foreach (DataRowView entAzInfo in entitaAzioneInformazione)
                        {
                            object entita = (entAzInfo["SiglaEntitaRif"] is DBNull ? entAzInfo["SiglaEntita"] : entAzInfo["SiglaEntitaRif"]);

                            Tuple<int, int>[] riga = nomiDefiniti[DefinedNames.GetName(entita, entAzInfo["SiglaInformazione"], suffissoData)];
                            Excel.Worksheet ws = CommonFunctions.WB.Sheets[nomeFoglio];
                            Excel.Range rng = ws.Range[ws.Cells[riga[0].Item1, riga[0].Item2], ws.Cells[riga[riga.Length - 1].Item1, riga[riga.Length - 1].Item2]];
                            object[,] tmpVal = rng.Value;
                            object[] values = tmpVal.Cast<object>().ToArray();

                            for (int i = 0, length = values.Length; i < length; i++)
                            {
                                DataRow row = dt.NewRow();

                                row["Campo1"] = nomeFoglio == "Iren Termo" ? "AHRP": "AIHRP";
                                row["Campo2"] = "Prod";
                                row["UP"] = codiceIF;
                                if (DefinedNames.IsDefined(nomeFoglio, DefinedNames.GetName(entita, "UNIT_COMM")))
                                    row["Campo3"] = "17";
                                else
                                    row["Campo3"] = "na";
                                row["Data"] = dataRif.Value.ToString("yyyy/MM/dd");
                                row["Ora"] = i + 1;
                                row["Informazione"] = entAzInfo["SiglaInformazione"].Equals("PMAX") ? "Pmax" : "Pmin";
                                row["Valore"] = values[i] ?? 0;

                                dt.Rows.Add(row);
                            }
                        }

                        var settings = (UserConfiguration)ConfigurationManager.GetSection("usrConfig");
                        var path = (UserConfigElement)settings.Items["pathExportMP_MGP"];

                        if (Directory.Exists(path.Value))
                        {
                            if (!ExportToCSV(System.IO.Path.Combine(path.Value, "AEM_" + (nomeFoglio == "Iren Termo" ? "AHRP_" : "AIHRP_") + codiceIF + "_" + dataRif.Value.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfffffff") + ".csv"), dt))
                                return false;
                        }
                        else
                        {
                            System.Windows.Forms.MessageBox.Show("Il percorso '" + path + "' non è raggiungibile.", Simboli.nomeApplicazione, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return false;   
                        }

                        break;
                }
                CommonFunctions.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, dataRif);
                CommonFunctions.DB.CloseConnection();
                return true;
            }
            catch (Exception e)
            {
                //TODO riabilitare log!!
                //InsertLog(DataBase.TipologiaLOG.LogErrore, "modProgram EsportaAzioneInformazione [" + siglaEntita + ", " + siglaAzione + "]: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                CommonFunctions.DB.CloseConnection();
                return false;
            }
        }

        private bool ExportToCSV(string nomeFile, DataTable dt)
        {
            try
            {
                using (StreamWriter outFile = new StreamWriter(nomeFile))
                {
                    foreach (DataRow r in dt.Rows)
                    {
                        IEnumerable<string> fields = r.ItemArray.Select(field => field.ToString());
                        outFile.WriteLine(string.Join(";", fields));
                    }
                    outFile.Flush();
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
