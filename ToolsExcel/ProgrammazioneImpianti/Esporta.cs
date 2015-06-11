﻿using Iren.ToolsExcel.UserConfig;
using Iren.ToolsExcel.Core;
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
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;


namespace Iren.ToolsExcel
{
    public class Esporta : AEsporta
    {
        public override bool RunExport(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            try
            {
                if (EsportaAzioneInformazione(siglaEntita, siglaAzione, desEntita, desAzione, dataRif))
                {
                    if(_db.OpenConnection())
                        Utility.DataBase.InsertApplicazioneRiepilogo(siglaEntita, siglaAzione, dataRif);
                    
                    _db.CloseConnection();
                    
                    return true;
                }

                return false;
            }
            catch (Exception e)
            {
                if (_db.OpenConnection())
                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "RunExport [" + siglaEntita + ", " + siglaAzione + "]: " + e.Message);

                _db.CloseConnection();

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
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
            entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            DefinedNames definedNames = new DefinedNames(nomeFoglio);

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

                    string suffissoData = Utility.Date.GetSuffissoData(dataRif);
                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        object siglaEntitaRif = (info["SiglaEntitaRif"] is DBNull ? info["SiglaEntita"] : info["SiglaEntitaRif"]);
                        
                        Excel.Worksheet ws = Workbook.WB.Sheets[nomeFoglio];
                        Range range = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], suffissoData);
                        range.Extend(0, definedNames.GetDayOffset(suffissoData) - 1);
                        Excel.Range rng = ws.Range[range.ToString()];

                        object[,] tmpVal = rng.Value;
                        object[] values = tmpVal.Cast<object>().ToArray();

                        for (int i = 0, length = values.Length; i < length; i++)
                        {
                            DataRow row = dt.NewRow();

                            row["Campo1"] = nomeFoglio == "Iren Termo" ? "AHRP" : "AIHRP";
                            row["Campo2"] = "Prod";
                            row["UP"] = codiceIF;
                            if (definedNames.IsDefined(siglaEntitaRif, "UNIT_COMM"))
                                row["Campo3"] = "17";
                            else
                                row["Campo3"] = "NA";
                            row["Data"] = dataRif.ToString("yyyy/MM/dd");
                            row["Ora"] = i + 1;
                            row["Informazione"] = info["SiglaInformazione"].Equals("PMAX") ? "Pmax" : "Pmin";
                            row["Valore"] = values[i] ?? 0;

                            dt.Rows.Add(row);
                        }
                    }

                    var path = Utility.Workbook.GetUsrConfigElement("pathExportMP_MGP");

                    string pathStr = Utility.ExportPath.PreparePath(path.Value);

                    if (Directory.Exists(pathStr))
                    {
                        if (!ExportToCSV(System.IO.Path.Combine(pathStr, "AEM_" + (nomeFoglio == "Iren Termo" ? "AHRP_" : "AIHRP_") + codiceIF + "_" + dataRif.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfffffff") + ".csv"), dt))
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