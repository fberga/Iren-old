using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.UserConfig;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    class Esporta : Base.Esporta
    {
        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            DataView entitaAzione = _localDB.Tables[Utility.DataBase.Tab.ENTITAAZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
            if (entitaAzione.Count == 0)
                return false;

            switch (siglaAzione.ToString())
            {
                case "E_VDT":
                    DataView entitaAssetto = _localDB.Tables[Utility.DataBase.Tab.ENTITAASSETTO].DefaultView;
                    entitaAssetto.RowFilter = "SiglaEntita = '" + siglaEntita + "'";

                    Dictionary<string,int> assettoFasce = new Dictionary<string,int>();
                    foreach (DataRowView assetto in entitaAssetto)
                        assettoFasce.Add((string)assetto["IdAssetto"], (int)assetto["NumeroFasce"]);

                    var path = Utility.Utilities.GetUsrConfigElement("pathExportSisComTerna");
                    string pathStr = Utility.ExportPath.PreparePath(path.Value);

                    

                    if (Directory.Exists(pathStr))
                    {
                        if (!CreaVariazioneDatiTecniciXML(siglaEntita, pathStr, assettoFasce))
                            return false;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Il percorso '" + pathStr + "' non è raggiungibile.", Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                        return false;
                    }
                    
                    break;
            }
            return true;
        }

        protected bool CreaVariazioneDatiTecniciXML(object siglaEntita, string exportPath, Dictionary<string,int> assettoFasce)
        {
            try
            {
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                Excel.Worksheet ws = Utility.Workbook.WB.Sheets[nomeFoglio];

                DateTime data = Utility.DataBase.DataAttiva;
                int oreData = Utility.Date.GetOreGiorno(Utility.DataBase.DataAttiva);

                DataView categoriaEntita = _localDB.Tables[Utility.DataBase.Tab.CATEGORIAENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];
                bool isTermo = categoriaEntita[0]["SiglaCategoria"].Equals("IREN_60T");

                XElement inserisci = new XElement("INSERISCI");

                DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

                Tuple<int, int>[] rangeEntita = nomiDefiniti.Get(DefinedNames.GetName(siglaEntita), "GOTO");

                IEnumerable<int> rows =
                    from value in rangeEntita
                    select value.Item1;

                IEnumerable<int> cols =
                    from value in rangeEntita
                    select value.Item2;

                Tuple<int, int> topLeft = new Tuple<int, int>(rows.Min(), cols.Min());
                Tuple<int, int> bottomRight = new Tuple<int, int>(rows.Max(), cols.Max());

                object[,] tmp = ws.Range[ws.Cells[topLeft.Item1, topLeft.Item2], ws.Cells[bottomRight.Item1, bottomRight.Item2]].Value;
                object[,] valoriEntita = new object[tmp.GetLength(0), tmp.GetLength(1)];
                Array.Copy(tmp, 1, valoriEntita, 0, valoriEntita.Length);

                Dictionary<string, Tuple<int, int>[]> ranges = new Dictionary<string, Tuple<int, int>[]>();

                int assetto = 1;
                foreach (KeyValuePair<string, int> assettoFascia in assettoFasce)
                {
                    ranges.Add(DefinedNames.GetName(siglaEntita, "TRISP_ASSETTO" + assetto),
                        nomiDefiniti[DefinedNames.GetName(siglaEntita, "TRISP_ASSETTO" + assetto)]);

                    ranges.Add(DefinedNames.GetName(siglaEntita, "GPA_ASSETTO" + assetto),
                        nomiDefiniti[DefinedNames.GetName(siglaEntita, "GPA_ASSETTO" + assetto)]);

                    ranges.Add(DefinedNames.GetName(siglaEntita, "GPD_ASSETTO" + assetto),
                        nomiDefiniti[DefinedNames.GetName(siglaEntita, "GPD_ASSETTO" + assetto)]);

                    ranges.Add(DefinedNames.GetName(siglaEntita, "TAVA_ASSETTO" + assetto),
                        nomiDefiniti[DefinedNames.GetName(siglaEntita, "TAVA_ASSETTO" + assetto)]);

                    ranges.Add(DefinedNames.GetName(siglaEntita, "TARA_ASSETTO" + assetto),
                        nomiDefiniti[DefinedNames.GetName(siglaEntita, "TARA_ASSETTO" + assetto)]);

                    ranges.Add(DefinedNames.GetName(siglaEntita, "BRS_ASSETTO" + assetto),
                        nomiDefiniti[DefinedNames.GetName(siglaEntita, "BRS_ASSETTO" + assetto)]);

                    if (isTermo)
                    {
                        ranges.Add(DefinedNames.GetName(siglaEntita, "TDERAMPA_ASSETTO" + assetto),
                            nomiDefiniti[DefinedNames.GetName(siglaEntita, "TDERAMPA_ASSETTO" + assetto)]);
                    }

                    for (int j = 1; j <= assettoFascia.Value; j++)
                    {
                        ranges.Add(DefinedNames.GetName(siglaEntita, "PSMIN_ASSETTO" + assetto + "_FASCIA" + j),
                            nomiDefiniti[DefinedNames.GetName(siglaEntita, "PSMIN_ASSETTO" + assetto + "_FASCIA" + j)]);
                        ranges.Add(DefinedNames.GetName(siglaEntita, "PSMIN_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j),
                            nomiDefiniti[DefinedNames.GetName(siglaEntita, "PSMIN_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j)]);

                        ranges.Add(DefinedNames.GetName(siglaEntita, "PSMAX_ASSETTO" + assetto + "_FASCIA" + j),
                            nomiDefiniti[DefinedNames.GetName(siglaEntita, "PSMAX_ASSETTO" + assetto + "_FASCIA" + j)]);
                        ranges.Add(DefinedNames.GetName(siglaEntita, "PSMAX_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j),
                            nomiDefiniti[DefinedNames.GetName(siglaEntita, "PSMAX_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j)]);

                        ranges.Add(DefinedNames.GetName(siglaEntita, "PTMIN_ASSETTO" + assetto + "_FASCIA" + j),
                        nomiDefiniti[DefinedNames.GetName(siglaEntita, "PTMIN_ASSETTO" + assetto + "_FASCIA" + j)]);

                        ranges.Add(DefinedNames.GetName(siglaEntita, "PTMAX_ASSETTO" + assetto + "_FASCIA" + j),
                            nomiDefiniti[DefinedNames.GetName(siglaEntita, "PTMAX_ASSETTO" + assetto + "_FASCIA" + j)]);
                    }

                    assetto++;
                }

                for (int j = 1; j <= oreData; j++)
                {
                    ranges.Add(DefinedNames.GetName(siglaEntita, "PQNR", Utility.Date.GetSuffissoData(data), "H" + j),
                        nomiDefiniti.GetByFilter(DefinedNames.Fields.Foglio + " = '" + nomeFoglio + "' AND " +
                                                 DefinedNames.Fields.Nome + " LIKE '" + DefinedNames.GetName(siglaEntita, "PQNR") + "%' AND " +
                                                 DefinedNames.Fields.Nome + " NOT LIKE '%PROFILO%' AND " +
                                                 DefinedNames.Fields.Nome + " LIKE '%" + DefinedNames.GetName(Utility.Date.GetSuffissoData(data)) + ".H" + j + "'"));
                }


                for (int i = 0; i < oreData && i < 24; i++)
                {
                    string start = data.ToString("yyyy-MM-dd") + "T" + i.ToString("00") + ":00:00";
                    string end = data.ToString("yyyy-MM-dd") + "T" + (i < 23 ? (i + 1).ToString("00") + ":00:00" : "23:59:00");

                    XElement vdt = new XElement("VDT", new XAttribute("DATAORAINIZIO", start), new XAttribute("DATAORAFINE", end),
                        new XElement("CODICEETSO", codiceRUP),
                        new XElement("IDMOTIVAZIONE", "VDT_VIN_TEC_UNI_PRO"),
                        new XElement("NOTE", "Vincoli Tecnologici dell'Unita di Produzione")
                    );

                    assetto = 1;
                    foreach (KeyValuePair<string, int> assettoFascia in assettoFasce)
                    {
                        for (int j = 1; j <= assettoFascia.Value; j++)
                        {
                            Tuple<int, int> cella = ranges[DefinedNames.GetName(siglaEntita, "PSMIN_ASSETTO" + assetto + "_FASCIA" + j)][i];
                            Tuple<int, int> cellaCorr = ranges[DefinedNames.GetName(siglaEntita, "PSMIN_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j)][i];
                            string psminVal = (valoriEntita[cellaCorr.Item1 - topLeft.Item1, cellaCorr.Item2 - topLeft.Item2] ?? valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2]).ToString().Replace('.', ',');

                            cella = ranges[DefinedNames.GetName(siglaEntita, "PSMAX_ASSETTO" + assetto + "_FASCIA" + j)][i];
                            cellaCorr = ranges[DefinedNames.GetName(siglaEntita, "PSMAX_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j)][i];
                            string psmaxVal = (valoriEntita[cellaCorr.Item1 - topLeft.Item1, cellaCorr.Item2 - topLeft.Item2] ?? valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2]).ToString().Replace('.', ',');

                            cella = ranges[DefinedNames.GetName(siglaEntita, "PTMIN_ASSETTO" + assetto + "_FASCIA" + j)][i];
                            string ptminVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2].ToString().Replace('.', ',');

                            cella = ranges[DefinedNames.GetName(siglaEntita, "PTMAX_ASSETTO" + assetto + "_FASCIA" + j)][i];
                            string ptmaxVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2].ToString().Replace('.', ',');

                            cella = ranges[DefinedNames.GetName(siglaEntita, "TRISP_ASSETTO" + assetto)][i];
                            string trispVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2].ToString().Replace('.', ',');

                            cella = ranges[DefinedNames.GetName(siglaEntita, "GPA_ASSETTO" + assetto)][i];
                            string gpaVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2].ToString().Replace('.', ',');

                            cella = ranges[DefinedNames.GetName(siglaEntita, "GPD_ASSETTO" + assetto)][i];
                            string gpdVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2].ToString().Replace('.', ',');

                            cella = ranges[DefinedNames.GetName(siglaEntita, "TAVA_ASSETTO" + assetto)][i];
                            string tavaVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2].ToString().Replace('.', ',');

                            cella = ranges[DefinedNames.GetName(siglaEntita, "TARA_ASSETTO" + assetto)][i];
                            string taraVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2].ToString().Replace('.', ',');

                            cella = ranges[DefinedNames.GetName(siglaEntita, "BRS_ASSETTO" + assetto)][i];
                            string brsVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2].ToString().Replace('.', ',');

                            string tderampaVal = null;
                            if (isTermo)
                            {
                                cella = ranges[DefinedNames.GetName(siglaEntita, "TDERAMPA_ASSETTO" + assetto)][i];
                                tderampaVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2].ToString().Replace('.', ',');
                            }

                            vdt.Add(new XElement("FASCIA",
                                new XElement("PSMIN", psminVal),
                                new XElement("PSMAX", psmaxVal),
                                new XElement("ASSETTO",
                                        new XElement("IDASSETTO", assettoFascia.Key),
                                        new XElement("PTMIN", ptminVal),
                                        new XElement("PTMAX", ptmaxVal),
                                        new XElement("TRISP", trispVal),
                                        new XElement("GPA", gpaVal),
                                        new XElement("GPD", gpdVal),
                                        new XElement("TAVA", tavaVal),
                                        new XElement("TARA", taraVal),
                                        new XElement("BRS", brsVal),
                                        (isTermo && tderampaVal != null ? new XElement("TDERAMPA", tderampaVal) : null)
                                    )
                                )
                            );
                        }
                        assetto++;
                    }
                    if (isTermo)
                    {
                        XElement pqnr = new XElement("PQNR");
                        for (int j = 0; j < 24; j++)
                        {
                            Tuple<int, int> cella = ranges[DefinedNames.GetName(siglaEntita, "PQNR", Utility.Date.GetSuffissoData(data), "H" + (i + 1))][j];
                            object pqnrVal = valoriEntita[cella.Item1 - topLeft.Item1, cella.Item2 - topLeft.Item2];
                            if (pqnrVal != null)
                                pqnr.Add(new XElement("Q", pqnrVal.ToString()));
                        }
                        vdt.Add(pqnr);
                    }

                    inserisci.Add(vdt);
                }

                XDocument variazioneDatiTecnici = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                    new XElement("FLUSSO", new XAttribute(XNamespace.Xmlns + "xsi", "http://www.w3.org/2001/XMLSchema-instance"),
                        inserisci)
                    );

                string filename = "VDT_" + codiceRUP.ToString().ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";
                variazioneDatiTecnici.Save(Path.Combine(exportPath, filename));

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
