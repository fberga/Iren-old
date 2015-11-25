﻿using Iren.PSO;
using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Funzione di esportazione custom.
    /// </summary>
    class Esporta : AEsporta
    {
        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            DataView entitaAzione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;
            if (entitaAzione.Count == 0)
                return false;

            switch (siglaAzione.ToString())
            {
                case "E_VDT":
                    DataView entitaAssetto = Workbook.Repository[DataBase.TAB.ENTITA_ASSETTO].DefaultView;
                    entitaAssetto.RowFilter = "SiglaEntita = '" + siglaEntita + "'";

                    Dictionary<string,int> assettoFasce = new Dictionary<string,int>();
                    foreach (DataRowView assetto in entitaAssetto)
                        assettoFasce.Add((string)assetto["IdAssetto"], (int)assetto["NumeroFasce"]);

                    string pathStr = PreparePath(Workbook.GetUsrConfigElement("pathExportSisComTerna"));

                    if (Directory.Exists(pathStr))
                    {
                        if (!CreaVariazioneDatiTecniciXML(siglaEntita, pathStr, assettoFasce))
                            return false;
                    }
                    else
                    {
                        System.Windows.Forms.MessageBox.Show("Il percorso '" + pathStr + "' non è raggiungibile.", Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                        return false;
                    }
                    
                    break;
                case "MAIL":
                    Workbook.ScreenUpdating = false;                    
                    string nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                    DefinedNames definedNames = new DefinedNames(nomeFoglio, DefinedNames.InitType.Naming);

                    var oldActiveWindow = Globals.ThisWorkbook.Application.ActiveWindow;
                    Globals.ThisWorkbook.Worksheets[nomeFoglio].Activate();

                    List<Range> export = new List<Range>();

                    //titolo entità
                    export.Add(new Range(definedNames.GetRowByNameSuffissoData(siglaEntita, "T", Date.SuffissoDATA1), definedNames.GetFirstCol() - 2).Extend(colOffset: 2 + Date.GetOreGiorno(Workbook.DataAttiva)));

                    //data
                    export.Add(new Range(Globals.ThisWorkbook.Application.ActiveWindow.SplitRow - 1, definedNames.GetFirstCol() - 2).Extend(colOffset: 2 + Date.GetOreGiorno(Workbook.DataAttiva)));

                    //ora
                    export.Add(new Range(Globals.ThisWorkbook.Application.ActiveWindow.SplitRow, definedNames.GetFirstCol() - 2).Extend(colOffset: 2 + Date.GetOreGiorno(Workbook.DataAttiva)));


                    DataView entitaAzioneInformazione = Workbook.Repository[DataBase.TAB.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                    entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Workbook.IdApplicazione;
                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        export.Add(new Range(definedNames.GetRowByNameSuffissoData(siglaEntita, info["SiglaInformazione"], Date.SuffissoDATA1), definedNames.GetFirstCol() - 2).Extend(colOffset: 2 + Date.GetOreGiorno(Workbook.DataAttiva)));
                    }

                    if (InviaMail(nomeFoglio, siglaEntita, export))
                    {

                    }

                    oldActiveWindow.Activate();

                    Workbook.ScreenUpdating = true;
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

                DateTime giorno = Workbook.DataAttiva;
                string suffissoData = Date.GetSuffissoData(giorno);
                int oreGiorno = Date.GetOreGiorno(Workbook.DataAttiva);

                DataView categoriaEntita = Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Workbook.IdApplicazione;
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
                        Range rngProfiloPQNR = definedNames.Get(siglaEntita, "PQNR_PROFILO", suffissoData);
                        if (ws.Range[rngProfiloPQNR.ToString()].Value == null)
                        {
                            SplashScreen.Close();
                            System.Windows.Forms.MessageBox.Show("Non è stato definito alcun profilo PQNR per l'UP " + siglaEntita + ": l'esportazione verrà interrotta per questa UP. Compilare il suo profilo per poter esportare.", Simboli.NomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            SplashScreen.Show();

                            return false;
                        }


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

                DataView entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_ALLEGATO_EXCEL' AND IdApplicazione = " + Workbook.IdApplicazione;
                if (entitaProprieta.Count > 0)
                {
                    fileName = @"D:\" + entitaProprieta[0]["Valore"] + "_VDT_" + Workbook.DataAttiva.ToString("yyyyMMdd") + ".xls";

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
                    Marshal.ReleaseComObject(wb);

                    var config = Workbook.GetUsrConfigElement("destMailTest");
                    string mailTo = config.Value;
                    string mailCC = "";

                    if (Workbook.Ambiente == "Produzione")
                    {
                        entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_MAIL_TO' AND IdApplicazione = " + Workbook.IdApplicazione;
                        mailTo = entitaProprieta[0]["Valore"].ToString();
                        entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_MAIL_CC' AND IdApplicazione = " + Workbook.IdApplicazione;
                        mailCC = entitaProprieta[0]["Valore"].ToString();
                    }

                    entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_CODICE_MAIL' AND IdApplicazione = " + Workbook.IdApplicazione;
                    string codUP = entitaProprieta[0]["Valore"].ToString();

                    config = Workbook.GetUsrConfigElement("oggettoMail");
                    string oggetto = config.Value.Replace("%COD%", codUP).Replace("%DATA%", Workbook.DataAttiva.ToString("dd-MM-yyyy"));
                    config = Workbook.GetUsrConfigElement("messaggioMail");
                    string messaggio = config.Value;
                    messaggio = Regex.Replace(messaggio, @"^[^\S\r\n]+", "", RegexOptions.Multiline);

                    Outlook.Application outlook = GetOutlookInstance();
                    Outlook._MailItem mail = outlook.CreateItem(Outlook.OlItemType.olMailItem);
                    
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

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                if(File.Exists(fileName))
                    File.Delete(fileName);

                return false;
            }

            return true;
        }
    }
}
