using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.UserConfig;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
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
                case "DATO_TOPICO":

                    var path = Utility.Workbook.GetUsrConfigElement("pathExportDatiTopici");
                    string pathStr = Utility.ExportPath.PreparePath(path.Value);

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
            }
            return true;
        }

        protected bool CreaDatiTopiciUnitaXML(object siglaEntita, object siglaAzione, string exportPath, DateTime dataRif)
        {
            try
            {
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                DefinedNames definedNames = new DefinedNames(nomeFoglio);
                Excel.Worksheet ws = Utility.Workbook.Sheets[nomeFoglio];

                string suffissoData = Utility.Date.GetSuffissoData(dataRif);
                int oreGiorno = Utility.Date.GetOreGiorno(dataRif);

                DataView categoriaEntita = _localDB.Tables[Utility.DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];
                //bool isTermo = categoriaEntita[0]["SiglaCategoria"].Equals("IREN_60T");

                DataView entitaAzioneInformazione = _localDB.Tables[Utility.DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

                XNamespace ns = XNamespace.Get("urn:XML-BIDMGM");

                XElement unit = new XElement(ns + "Unit", new XAttribute("StartDate", dataRif.ToString("yyyyMMdd")), new XAttribute("IDUnit", codiceRUP));

                for (int i = 0; i < oreGiorno; i++)
                {
                    string[] values = new string[7];
                    int j = 0;
                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"];
                        Range rng = definedNames.Get(siglaEntitaRif, info["SiglaInformazione"], suffissoData, Utility.Date.GetSuffissoOra(i + 1));
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
    }
}
