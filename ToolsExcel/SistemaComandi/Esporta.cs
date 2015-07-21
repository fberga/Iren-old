using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.UserConfig;
using Iren.ToolsExcel.Utility;
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
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Iren.ToolsExcel
{
    class Esporta : AEsporta
    {
        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            DataView entitaAzione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
            if (entitaAzione.Count == 0)
                return false;

            switch (siglaAzione.ToString())
            {
                case "E_VDT":
                    DataView entitaAssetto = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_ASSETTO].DefaultView;
                    entitaAssetto.RowFilter = "SiglaEntita = '" + siglaEntita + "'";

                    Dictionary<string,int> assettoFasce = new Dictionary<string,int>();
                    foreach (DataRowView assetto in entitaAssetto)
                        assettoFasce.Add((string)assetto["IdAssetto"], (int)assetto["NumeroFasce"]);

                    var path = Workbook.GetUsrConfigElement("pathExportSisComTerna");
                    string pathStr = PreparePath(path.Value);

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
                    Globals.ThisWorkbook.Application.ScreenUpdating = false;
                    string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                    DefinedNames definedNames = new DefinedNames(nomeFoglio, DefinedNames.InitType.Naming);

                    var oldActiveWindow = Globals.ThisWorkbook.Application.ActiveWindow;
                    Globals.ThisWorkbook.Worksheets[nomeFoglio].Activate();

                    List<Range> export = new List<Range>();

                    //titolo entità
                    export.Add(new Range(definedNames.GetRowByNameSuffissoData(siglaEntita, "T", Date.SuffissoDATA1), definedNames.GetFirstCol() - 2).Extend(colOffset: 2 + Date.GetOreGiorno(DataBase.DataAttiva)));

                    //data
                    export.Add(new Range(Globals.ThisWorkbook.Application.ActiveWindow.SplitRow - 1, definedNames.GetFirstCol() - 2).Extend(colOffset: 2 + Date.GetOreGiorno(DataBase.DataAttiva)));

                    //ora
                    export.Add(new Range(Globals.ThisWorkbook.Application.ActiveWindow.SplitRow, definedNames.GetFirstCol() - 2).Extend(colOffset: 2 + Date.GetOreGiorno(DataBase.DataAttiva)));


                    DataView entitaAzioneInformazione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                    entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        export.Add(new Range(definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], Date.SuffissoDATA1), definedNames.GetFirstCol() - 2).Extend(colOffset: 2 + Date.GetOreGiorno(DataBase.DataAttiva)));
                    }

                    if (InviaMail(nomeFoglio, siglaEntita, export))
                    {

                    }

                    oldActiveWindow.Activate();



                    Globals.ThisWorkbook.Application.ScreenUpdating = true;
                    break;
            }
            return true;
        }

        protected bool CreaVariazioneDatiTecniciXML(object siglaEntita, string exportPath, Dictionary<string,int> assettoFasce)
        {
            try
            {
                string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                DefinedNames definedNames = new DefinedNames(nomeFoglio);
                Excel.Worksheet ws = Workbook.Sheets[nomeFoglio];

                DateTime giorno = DataBase.DataAttiva;
                string suffissoData = Date.GetSuffissoData(giorno);
                int oreGiorno = Date.GetOreGiorno(DataBase.DataAttiva);

                DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
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
                            Range rng = definedNames.Get(siglaEntita, "PSMIN_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Date.GetSuffissoOra(i + 1));
                            Range rngCorr = definedNames.Get(siglaEntita, "PSMIN_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Date.GetSuffissoOra(i + 1));
                            string psminVal = (ws.Range[rngCorr.ToString()].Value ?? ws.Range[rng.ToString()].Value).ToString().Replace('.', ',');

                            rng = definedNames.Get(siglaEntita, "PSMAX_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Date.GetSuffissoOra(i + 1));
                            rngCorr = definedNames.Get(siglaEntita, "PSMAX_CORRETTA_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Date.GetSuffissoOra(i + 1));
                            string psmaxVal = (ws.Range[rngCorr.ToString()].Value ?? ws.Range[rng.ToString()].Value).ToString().Replace('.', ',');

                            rng = definedNames.Get(siglaEntita, "PTMIN_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Date.GetSuffissoOra(i + 1));
                            string ptminVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = definedNames.Get(siglaEntita, "PTMAX_ASSETTO" + assetto + "_FASCIA" + j, suffissoData, Date.GetSuffissoOra(i + 1));
                            string ptmaxVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = definedNames.Get(siglaEntita, "TRISP_ASSETTO" + assetto, suffissoData, Date.GetSuffissoOra(i + 1));
                            string trispVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = definedNames.Get(siglaEntita, "GPA_ASSETTO" + assetto, suffissoData, Date.GetSuffissoOra(i + 1));
                            string gpaVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = definedNames.Get(siglaEntita, "GPD_ASSETTO" + assetto, suffissoData, Date.GetSuffissoOra(i + 1));
                            string gpdVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = definedNames.Get(siglaEntita, "TAVA_ASSETTO" + assetto, suffissoData, Date.GetSuffissoOra(i + 1));
                            string tavaVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = definedNames.Get(siglaEntita, "TARA_ASSETTO" + assetto, suffissoData, Date.GetSuffissoOra(i + 1));
                            string taraVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            rng = definedNames.Get(siglaEntita, "BRS_ASSETTO" + assetto, suffissoData, Date.GetSuffissoOra(i + 1));
                            string brsVal = ws.Range[rng.ToString()].Value.ToString().Replace('.', ',');

                            string tderampaVal = null;
                            if (isTermo)
                            {
                                rng = definedNames.Get(siglaEntita, "TDERAMPA_ASSETTO" + assetto, suffissoData, Date.GetSuffissoOra(i + 1));
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
                            Range rng = definedNames.Get(siglaEntita, "PQNR" + j, suffissoData, Date.GetSuffissoOra(i + 1));
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

        protected bool InviaMail(string nomeFoglio, object siglaEntita, List<Range> export) 
        {
            string fileName = "";
            try
            {
                Excel.Worksheet ws = Globals.ThisWorkbook.Sheets[nomeFoglio];

                DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_ALLEGATO_EXCEL'";
                if (entitaProprieta.Count > 0)
                {
                    fileName = @"D:\" + entitaProprieta[0]["Valore"] + "_VDT_" + DataBase.DataAttiva.ToString("yyyyMMdd") + ".xls";

                    Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Add();
                    int i = 2;
                    foreach (Range rng in export)
                    {
                        ws.Range[rng.ToString()].Copy();
                        wb.Sheets[1].Range["B" + i++].PasteSpecial();
                    }
                    wb.Sheets[1].Columns["B:C"].EntireColumn.AutoFit();
                    wb.Sheets[1].Range["A1"].Select();
                    wb.SaveAs(fileName, Excel.XlFileFormat.xlExcel8);
                    wb.Close();

                    var config = Workbook.GetUsrConfigElement("destMailTest");
                    string mailTo = config.Value;
                    string mailCC = "";

                    if (Simboli.Ambiente == "Produzione")
                    {
                        entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_MAIL_TO'";
                        mailTo = entitaProprieta[0]["Valore"].ToString();
                        entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_MAIL_CC'";
                        mailCC = entitaProprieta[0]["Valore"].ToString();
                    }
                    
                    entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_CODICE_MAIL'";
                    string codUP = entitaProprieta[0]["Valore"].ToString();

                    config = Workbook.GetUsrConfigElement("oggettoMail");
                    string oggetto = config.Value.Replace("%COD%", codUP).Replace("%DATA%", DataBase.DataAttiva.ToString("dd-MM-yyyy"));
                    config = Workbook.GetUsrConfigElement("messaggioMail");
                    string messaggio = config.Value;
                    messaggio = Regex.Replace(messaggio, @"^[^\S\r\n]+", "", RegexOptions.Multiline);

                    Outlook.Application outlook = GetOutlookInstance();
                    Outlook.MailItem mail = outlook.CreateItem(Outlook.OlItemType.olMailItem);
                    
                    //TODO check se manda sempre con lo stesso account...
                    Outlook.Account senderAccount = outlook.Session.Accounts[1];
                    foreach (Outlook.Account account in outlook.Session.Accounts)
                    {
                        if (account.DisplayName == "Bidding")
                            senderAccount = account;
                    }
                    mail.SendUsingAccount = senderAccount;
                    mail.Subject = oggetto;
                    mail.Body = messaggio;
                    mail.Recipients.Add(mailTo);
                    mail.CC = mailCC;
                    mail.Attachments.Add(fileName);

                    mail.Send();

                    File.Delete(fileName);
                }
            }
            catch(Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "SisCom - Esporta.InvioMail: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                if(File.Exists(fileName))
                    File.Delete(fileName);

                return false;
            }

            return true;
        }
    }
}
