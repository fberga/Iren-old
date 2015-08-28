using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.UserConfig;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    /// <summary>
    /// Funzioni di esportazione personalizzate.
    /// </summary>
    class Esporta : Base.Esporta
    {
        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            DataView entitaAzione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Simboli.AppID;
            if (entitaAzione.Count == 0)
                return false;

            UserConfigElement path;
            string pathStr;

            switch (siglaAzione.ToString())
            {
                case "DATO_TOPICO":

                    path = Workbook.GetUsrConfigElement("pathExportDatiTopici");
                    pathStr = PreparePath(path.Value);

                    if (Directory.Exists(pathStr))
                    {
                        if (!CreaDatiTopiciUnitaXML(siglaEntita, siglaAzione, pathStr, dataRif))
                            return false;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Il percorso '" + pathStr + "' non è raggiungibile.", Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                        return false;
                    }
                    
                    break;
                case "E_OFFERTA_MGP":

                    path = Workbook.GetUsrConfigElement("pathExportOFFERTE_MGP_GME");
                    pathStr = PreparePath(path.Value);

                    if (Directory.Exists(pathStr))
                    {
                        if (!CreaOfferteSuggeriteXML_GME(siglaEntita, siglaAzione, pathStr, dataRif))
                            return false;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Il percorso '" + pathStr + "' non è raggiungibile.", Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                        return false;
                    }

                    path = Workbook.GetUsrConfigElement("pathExportOFFERTE_MGP");
                    pathStr = PreparePath(path.Value);

                    if (Directory.Exists(pathStr))
                    {
                        if (!CreaOfferteSuggeriteXML(siglaEntita, siglaAzione, pathStr, dataRif))
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

        protected bool CreaDatiTopiciUnitaXML(object siglaEntita, object siglaAzione, string exportPath, DateTime dataRif)
        {
            try
            {
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                DefinedNames definedNames = new DefinedNames(nomeFoglio);
                Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

                string suffissoData = Date.GetSuffissoData(dataRif);
                int oreGiorno = Date.GetOreGiorno(dataRif);

                DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Simboli.AppID;
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];
                //bool isTermo = categoriaEntita[0]["SiglaCategoria"].Equals("IREN_60T");

                DataView entitaAzioneInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Simboli.AppID;

                XNamespace ns = XNamespace.Get("urn:XML-BIDMGM");

                XElement unit = new XElement(ns + "Unit", new XAttribute("StartDate", dataRif.ToString("yyyyMMdd")), new XAttribute("IDUnit", codiceRUP));

                for (int i = 0; i < oreGiorno; i++)
                {
                    string[] values = new string[7];
                    int j = 0;
                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"];
                        Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], suffissoData, Date.GetSuffissoOra(i + 1));
                        values[j++] = (ws.Range[rng.ToString()].Value ?? "0").ToString().Replace('.', ',');

                    }

                    unit.Add(
                        new XElement(ns + "PR", i + 1,
                            new XAttribute("OPTIMAL", values[0] ?? "0"),
                            new XAttribute("MaxPower", values[1] ?? "0"),
                            new XAttribute("MinTech", values[2] ?? "0"),
                            new XAttribute("ReqPow", values[3] ?? "0"),
                            new XAttribute("COST", values[4] ?? "0"),
                            new XAttribute("COST2", values[5] ?? "0"),
                            new XAttribute("PumpingPower", values[6] ?? "0")
                        )
                    );
                }

                XNamespace xsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
                XNamespace schemaLocation = XNamespace.Get("urn:XML-BIDMGM BM_DatiTopiciUnita.xsd");

                XDocument datiTopiciUnita = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                    new XElement(ns + "BMTransaction-DTU",
                            new XAttribute("ReferenceNumber", codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss")), 
                            new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                            new XAttribute(xsi + "schemaLocation", schemaLocation), 
                            new XElement(ns + "DatiTopiciUnit", 
                                unit))
                    );

                string filename = "DatiTopici_" + codiceRUP.ToString().ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";
                datiTopiciUnita.Save(Path.Combine(exportPath, filename));

                return true;
            }
            catch
            {
                return false;
            }
        }
        protected bool CreaOfferteSuggeriteXML_GME(object siglaEntita, object siglaAzione, string exportPath, DateTime dataRif)
        {
            try
            {
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                DefinedNames definedNames = new DefinedNames(nomeFoglio);
                Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

                string suffissoData = Date.GetSuffissoData(dataRif);
                int oreGiorno = Date.GetOreGiorno(dataRif);

                DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Simboli.AppID;
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];

                DataView entitaAzioneInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione ='" + siglaAzione + "' AND SiglaInformazione LIKE 'OFFERTA_MGP_E%' AND IdApplicazione = " + Simboli.AppID;

                DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'OFFERTA_MGP_TIPO_OFFERTA' AND IdApplicazione = " + Simboli.AppID;

                XNamespace ns = XNamespace.Get("urn:XML-PIPE");
                XNamespace xsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
                XNamespace xsd = XNamespace.Get("http://www.w3.org/2001/XMLSchema");
                XNamespace schemaLocation = XNamespace.Get("urn:XML-PIPE PIPEDocument.xsd");

                string referenceNumber = codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                XElement PIPEDocument = new XElement(ns + "PIPEDocument",
                        new XAttribute("ReferenceNumber", referenceNumber.Length > 30 ? referenceNumber.Substring(0, 30) : referenceNumber),
                        new XAttribute("CreationDate", DateTime.Now.ToString("yyyyMMddHHmmss")),
                        new XAttribute("Version", "1.0"),
                        new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                        new XAttribute(XNamespace.Xmlns + "xsd", xsd),
                        new XAttribute(xsi + "schemaLocation", schemaLocation),
                        new XElement(ns + "TradingPartnerDirectory",
                            new XElement(ns + "Sender",
                                new XElement(ns + "TradingPartner",
                                    new XAttribute("PartnerType", "Market Participant"),
                                    new XElement(ns + "CompanyName", "IREN MERCATO S.P.A."),
                                    new XElement(ns + "CompanyIdentifier", "OEACSMG")
                                )
                            ),
                            new XElement(ns + "Recipient",
                                new XElement(ns + "TradingPartner",
                                    new XAttribute("PartnerType", "Operator"),
                                    new XElement(ns + "CompanyName", "GME SPA"),
                                    new XElement(ns + "CompanyIdentifier", "IDGME")
                                )
                            )
                        )
                    );

                foreach(DataRowView info in entitaAzioneInformazione)
                {
                    string gradino = Regex.Match(info["SiglaInformazione"].ToString(), @"\d+").Value;
                    object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"];
                    Range rngEnergia = definedNames.Get(siglaEntitaRif, "OFFERTA_MGP_E" + gradino).Extend(colOffset: oreGiorno);
                    Range rngPrezzo = definedNames.Get(siglaEntitaRif, "OFFERTA_MGP_P" + gradino).Extend(colOffset: oreGiorno);

                    for (int i = 0; i < oreGiorno; i++)
                    {
                        decimal valoreOfferta = Math.Abs((decimal)(ws.Range[rngEnergia.Columns[i].ToString()].Value ?? 0));
                        decimal prezzoOfferta = (decimal)(ws.Range[rngPrezzo.Columns[i].ToString()].Value ?? 0);

                        if (valoreOfferta != 0)
                        {
                            object tipoOfferta = entitaProprieta[0]["Valore"].Equals("MISTA") ? (valoreOfferta < 0 ? "ACQ" : "VEN") : entitaProprieta[0]["Valore"];

                            XElement bidSubmittal = new XElement(ns + "BidSubmittal",
                                    new XAttribute("MarketParticipantNumber", codiceRUP + "_" + dataRif.ToString("yyyyMMdd") + "_" + (i + 1) + "_G" + gradino),
                                    new XAttribute("ReplacementIndicator", "Yes"),
                                    new XAttribute("PredefinedOffer", "No"),
                                    new XAttribute("Purpose", tipoOfferta.Equals("VEN") ? "Sell" : "Buy"),
                                    new XElement(ns + "Market", "MGP"),
                                    new XElement(ns + "Date", dataRif.ToString("yyyyMMdd")),
                                    new XElement(ns + "Hour", i + 1),
                                    new XElement(ns + "UnitReferenceNumber", codiceRUP),
                                    new XElement(ns + "BidQuantity",
                                        new XAttribute("UnitOfMeasure", "MWh"), valoreOfferta.ToString(CultureInfo.InstalledUICulture)),
                                    new XElement(ns + "EnergyPrice", prezzoOfferta.ToString(CultureInfo.InstalledUICulture))
                                );

                            PIPEDocument.Add(new XElement(ns + "PIPTransaction", bidSubmittal));
                        }
                    }
                }

                XDocument offerteSuggerite = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                        PIPEDocument
                    );

                string filename = "Suggerite_MGP_" + codiceRUP.ToString() + "_GME.xml";
                offerteSuggerite.Save(Path.Combine(exportPath, filename));

                return true;
            }
            catch
            {
                return false;
            }
        }
        protected bool CreaOfferteSuggeriteXML(object siglaEntita, object siglaAzione, string exportPath, DateTime dataRif)
        {
            try
            {
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                DefinedNames definedNames = new DefinedNames(nomeFoglio);
                Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

                string suffissoData = Date.GetSuffissoData(dataRif);
                int oreGiorno = Date.GetOreGiorno(dataRif);

                DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Simboli.AppID;
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];

                DataView entitaAzioneInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione ='" + siglaAzione + "' AND SiglaInformazione LIKE 'OFFERTA_MGP_E%' AND IdApplicazione = " + Simboli.AppID;

                DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'OFFERTA_MGP_TIPO_OFFERTA' AND IdApplicazione = " + Simboli.AppID;

                XNamespace ns = XNamespace.Get("urn:XML-BIDMGM");
                XNamespace xsi = XNamespace.Get("http://www.w3.org/2001/XMLSchema-instance");
                XNamespace schemaLocation = XNamespace.Get("urn:XML-BIDMGM BM_SuggestedOffer.xsd");

                string referenceNumber = codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

                XElement BMTransaction = new XElement(ns + "BMTransaction-SUG",
                        new XAttribute("ReferenceNumber", referenceNumber.Length > 30 ? referenceNumber.Substring(0, 30) : referenceNumber),
                        new XAttribute(XNamespace.Xmlns + "xsi", xsi),
                        new XAttribute(xsi + "schemaLocation", schemaLocation),
                        new XElement(ns + "Suggested",
                            new XElement(ns + "Coordinate", 
                                new XAttribute("Mercato", "MGP"),
                                new XAttribute("IDUnit", codiceRUP),
                                new XAttribute("FlowDate", dataRif.ToString("yyyyMMdd"))
                            )
                        )
                    );

                foreach (DataRowView info in entitaAzioneInformazione)
                {
                    string gradino = Regex.Match(info["SiglaInformazione"].ToString(), @"\d+").Value;
                    object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"];
                    Range rngEnergia = definedNames.Get(siglaEntitaRif, "OFFERTA_MGP_E" + gradino).Extend(colOffset: oreGiorno);
                    Range rngPrezzo = definedNames.Get(siglaEntitaRif, "OFFERTA_MGP_P" + gradino).Extend(colOffset: oreGiorno);

                    for (int i = 0; i < oreGiorno; i++)
                    {
                        decimal valoreOfferta = Math.Abs((decimal)(ws.Range[rngEnergia.Columns[i].ToString()].Value ?? 0));
                        decimal prezzoOfferta = (decimal)(ws.Range[rngPrezzo.Columns[i].ToString()].Value ?? 0);

                        object tipoOfferta = entitaProprieta[0]["Valore"].Equals("MISTA") ? (valoreOfferta < 0 ? "ACQ" : "VEN") : entitaProprieta[0]["Valore"];

                        XElement sg = new XElement(ns + ("SG" + gradino),
                                new XAttribute("PRE", prezzoOfferta.ToString(CultureInfo.InstalledUICulture)),
                                new XAttribute("QUA", valoreOfferta.ToString(CultureInfo.InstalledUICulture)),
                                new XAttribute("AZIONE", tipoOfferta),
                                (i + 1)
                            );

                        BMTransaction.Element(ns + "Suggested").Element(ns + "Coordinate").Add(sg);
                    }
                }

                XDocument offerteSuggerite = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                        BMTransaction
                    );

                string filename = "Suggerite_MGP_" + codiceRUP.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";
                offerteSuggerite.Save(Path.Combine(exportPath, filename));

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
