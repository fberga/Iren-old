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
            DataView entitaAzione = _localDB.Tables[Utility.DataBase.Tab.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
            if (entitaAzione.Count == 0)
                return false;

            switch (siglaAzione.ToString())
            {
                case "E_VDT":
                    DataView entitaAssetto = _localDB.Tables[Utility.DataBase.Tab.ENTITA_ASSETTO].DefaultView;
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
                case "MAIL":


                    break;
            }
            return true;
        }

        protected bool CreaVariazioneDatiTecniciXML(object siglaEntita, string exportPath, Dictionary<string,int> assettoFasce)
        {
            try
            {
                string nomeFoglio = NewDefinedNames.GetSheetName(siglaEntita);
                NewDefinedNames newNomiDefiniti = new NewDefinedNames(nomeFoglio);
                Excel.Worksheet ws = Utility.Workbook.WB.Sheets[nomeFoglio];

                DateTime giorno = Utility.DataBase.DataAttiva;
                string suffissoData = Utility.Date.GetSuffissoData(giorno);
                int oreGiorno = Utility.Date.GetOreGiorno(Utility.DataBase.DataAttiva);

                DataView categoriaEntita = _localDB.Tables[Utility.DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];
                bool isTermo = categoriaEntita[0]["SiglaCategoria"].Equals("IREN_60T");

                XElement inserisci = new XElement("INSERISCI");

                
                for (int i = 0; i < oreGiorno && i < 24; i++)
                {
                    string start = giorno.ToString("yyyy-MM-dd") + "T" + i.ToString("00") + ":00:00";
                    string end = giorno.ToString("yyyy-MM-dd") + "T" + (i < 23 ? (i + 1).ToString("00") + ":00:00" : "23:59:00");

                    XElement vdt = new XElement("VDT", new XAttribute("DATAORAINIZIO", start), new XAttribute("DATAORAFINE", end),
                        new XElement("CODICEETSO", codiceRUP),
                        new XElement("IDMOTIVAZIONE", "VDT_VIN_TEC_UNI_PRO"),
                        new XElement("NOTE", "Vincoli Tecnologici dell'Unita di Produzione")
                    );

                    int assetto = 1;
                    foreach (KeyValuePair<string, int> assettoFascia in assettoFasce)
                    {
                        for (int j = 1; j <= assettoFascia.Value; j++)
                        {
                            Range rng = newNomiDefiniti.Get(siglaEntita, "PSMIN_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            Range rngCorr = newNomiDefiniti.Get(siglaEntita, "PSMIN_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string psminVal = (ws.Range[rngCorr.ToString()].Value ?? ws.Range[rng.ToString()].Value).ToString().Replace('.', ',');

                            rng = newNomiDefiniti.Get(siglaEntita, "PSMAX_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            rngCorr = newNomiDefiniti.Get(siglaEntita, "PSMAX_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string psmaxVal = (ws.Range[rngCorr.ToString()].Value ?? ws.Range[rng.ToString()].Value).ToString().Replace('.', ',');

                            rng = newNomiDefiniti.Get(siglaEntita, "PTMIN_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string ptminVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = newNomiDefiniti.Get(siglaEntita, "PTMAX_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string ptmaxVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = newNomiDefiniti.Get(siglaEntita, "TRISP_ASSETTO" + assetto, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string trispVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = newNomiDefiniti.Get(siglaEntita, "GPA_ASSETTO" + assetto, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string gpaVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = newNomiDefiniti.Get(siglaEntita, "GPD_ASSETTO" + assetto, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string gpdVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = newNomiDefiniti.Get(siglaEntita, "TAVA_ASSETTO" + assetto, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string tavaVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = newNomiDefiniti.Get(siglaEntita, "TARA_ASSETTO" + assetto, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string taraVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = newNomiDefiniti.Get(siglaEntita, "BRS_ASSETTO" + assetto, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            string brsVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            string tderampaVal = null;
                            if (isTermo)
                            {
                                rng = newNomiDefiniti.Get(siglaEntita, "TDERAMPA_ASSETTO" + assetto, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                                tderampaVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');
                            }
                            if (ptminVal != "" && ptmaxVal != "")
                            {
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
                        }
                        assetto++;
                    }
                    if (isTermo)
                    {
                        XElement pqnr = new XElement("PQNR");
                        for (int j = 1; j <= 24; j++)
                        {
                            Range rng = newNomiDefiniti.Get(siglaEntita, "PQNR" + j, suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                            object pqnrVal = ws.Range[rng.ToString()].Value;
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
