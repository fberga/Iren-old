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

                    string filename = "VDT_" + siglaEntita.ToString().ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                    if (Directory.Exists(pathStr) && CreaVariazioneDatiTecniciXML(siglaEntita, Path.Combine(pathStr, filename), assettoFasce))
                    { 
                        
                    }

                    break;

            }


            return true;
        }

        protected bool CreaVariazioneDatiTecniciXML(object siglaEntita, string exportPath, Dictionary<string,int> assettoFasce)
        {
            string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
            Excel.Worksheet ws = Utility.Workbook.WB.Sheets[nomeFoglio];
            
            DateTime data = Utility.DataBase.DataAttiva;
            int oreData = Utility.Date.GetOreGiorno(Utility.DataBase.DataAttiva);

            DataView categoriaEntita = _localDB.Tables[Utility.DataBase.Tab.CATEGORIAENTITA].DefaultView;
            categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
            object codiceRUP = categoriaEntita[0]["CodiceRUP"];
            bool isTermo = categoriaEntita[0]["SiglaCategoria"].Equals("IREN_60T");

            //comincio a creare l'XML
            XDocument variazioneDatiTecnici = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"));
            XElement flusso = new XElement("FLUSSO", new XAttribute(XNamespace.Xmlns + "xsi", "http://www.w3.org/2001/XMLSchema-instance"));
            XElement inserisci = new XElement("INSERISCI");
            
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);

            for (int i = 0; i < oreData; i++)
            {
                string start = data.ToString("yyyy-MM-dd") + "T" + i.ToString("00") + ":00:00";
                string end = data.ToString("yyyy-MM-dd") + "T" + (i < 23 ? (i + 1).ToString("00") + ":00:00" : "23:59:00");

                XElement vdt = new XElement("VDT", new XAttribute("DATAORAINIZIO", start), new XAttribute("DATAORAFINE", end),
                    new XElement("CODICEETSO", codiceRUP),
                    new XElement("IDMOTIVAZIONE", "VDT_VIN_TEC_UNI_PRO"),
                    new XElement("NOTE", "Vincoli Tecnologici dell'Unita di Produzione")
                );

                int assetto = 1;
                foreach (KeyValuePair<string,int> assettoFascia in assettoFasce)
                {

                    for (int j = 1; j <= assettoFascia.Value; j++)
                    {
                        Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "PSMIN_ASSETTO" + assetto + "_FASCIA" + j, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        Tuple<int, int> cellaCorr = nomiDefiniti[DefinedNames.GetName(siglaEntita, "PSMIN_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string psmin = (ws.Cells[cellaCorr.Item1, cellaCorr.Item2].Value ?? ws.Cells[cella.Item1, cella.Item2].Value).ToString().Replace('.',',');

                        cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "PSMAX_ASSETTO" + assetto + "_FASCIA" + j, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        cellaCorr = nomiDefiniti[DefinedNames.GetName(siglaEntita, "PSMAX_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string psmax = (ws.Cells[cellaCorr.Item1, cellaCorr.Item2].Value ?? ws.Cells[cella.Item1, cella.Item2].Value).ToString().Replace('.',',');

                        cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "PTMIN_ASSETTO" + assetto + "_FASCIA" + j, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string ptmin = ws.Cells[cella.Item1, cella.Item2].Value.ToString().Replace('.',',');

                        cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "PTMAX_ASSETTO" + assetto + "_FASCIA" + j, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string ptmax = ws.Cells[cella.Item1, cella.Item2].Value.ToString().Replace('.',',');

                        cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "TRISP_ASSETTO" + assetto, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string trisp = ws.Cells[cella.Item1, cella.Item2].Value.ToString().Replace('.',',');

                        cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "GPA_ASSETTO" + assetto, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string gpa = ws.Cells[cella.Item1, cella.Item2].Value.ToString().Replace('.',',');

                        cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "GPD_ASSETTO" + assetto, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string gpd = ws.Cells[cella.Item1, cella.Item2].Value.ToString().Replace('.',',');

                        cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "TAVA_ASSETTO" + assetto, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string tava = ws.Cells[cella.Item1, cella.Item2].Value.ToString().Replace('.',',');

                        cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "TARA_ASSETTO" + assetto, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string tara = ws.Cells[cella.Item1, cella.Item2].Value.ToString().Replace('.',',');

                        cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "BRS_ASSETTO" + assetto, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        string brs = ws.Cells[cella.Item1, cella.Item2].Value.ToString().Replace('.',',');

                        string tderampa = null;
                        if(isTermo)
                        {
                            cella = nomiDefiniti[DefinedNames.GetName(siglaEntita, "TDERAMPA_ASSETTO" + assetto, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                            tderampa = ws.Cells[cella.Item1, cella.Item2].Value.ToString().Replace('.',',');
                        }

                        vdt.Add(new XElement("FASCIA", 
                            new XElement("PSMIN", psmin),
                            new XElement("PSMAX", psmax),
                            new XElement("ASSETTO", 
                                    new XElement("IDASSETTO", assettoFascia.Key),
                                    new XElement("PTMIN", ptmin),
                                    new XElement("PTMAX", ptmax),
                                    new XElement("TRISP", trisp),
                                    new XElement("GPA", gpa),
                                    new XElement("GPD", gpd),
                                    new XElement("TAVA", tava),
                                    new XElement("TARA", tara),
                                    new XElement("BRS", brs),
                                    (isTermo ? new XElement("TDERAMPA", tderampa) : null)
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
                        Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName(codiceRUP, "PQNR" + j, Utility.Date.GetSuffissoData(data), "H" + (i + 1))][0];
                        object pqnrVal = ws.Cells[cella.Item1, cella.Item2].Value;
                        if (pqnrVal != null)
                            pqnr.Add(new XElement("Q", pqnrVal.ToString()));
                    }
                }


                inserisci.Add(vdt);
            }
            flusso.Add(inserisci);
            variazioneDatiTecnici.Add(flusso);

            variazioneDatiTecnici.Save(exportPath);

            //XDocument variazioneDatiTecnici = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
            //    new XElement(xmlns + "FLUSSO", )
                
                
            //    );



            return true;
        }
    }
}
