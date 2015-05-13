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
//using Outlook = Microsoft.Office.Interop.Outlook;

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

                    var path = Utility.Utilities.GetUsrConfigElement("pathExportDatiTopici");
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
                case "MAIL":
                    Globals.ThisWorkbook.Application.ScreenUpdating = false;
                    string nomeFoglio = NewDefinedNames.GetSheetName(siglaEntita);
                    NewDefinedNames newNomiDefiniti = new NewDefinedNames(nomeFoglio, NewDefinedNames.InitType.NamingOnly);

                    var oldActiveWindow = Globals.ThisWorkbook.Application.ActiveWindow;
                    Globals.ThisWorkbook.Worksheets[nomeFoglio].Activate();

                    List<Range> export = new List<Range>();

                    //titolo entità
                    export.Add(new Range(newNomiDefiniti.GetRowByName(siglaEntita, "T", Utility.Date.GetSuffissoDATA1), newNomiDefiniti.GetFirstCol() - 2).Extend(colOffset: 2 + Utility.Date.GetOreGiorno(Utility.DataBase.DataAttiva)));

                    //data
                    export.Add(new Range(Globals.ThisWorkbook.Application.ActiveWindow.SplitRow - 1, newNomiDefiniti.GetFirstCol() - 2).Extend(colOffset: 2 + Utility.Date.GetOreGiorno(Utility.DataBase.DataAttiva)));

                    //ora
                    export.Add(new Range(Globals.ThisWorkbook.Application.ActiveWindow.SplitRow, newNomiDefiniti.GetFirstCol() - 2).Extend(colOffset: 2 + Utility.Date.GetOreGiorno(Utility.DataBase.DataAttiva)));


                    DataView entitaAzioneInformazione = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                    entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        export.Add(new Range(newNomiDefiniti.GetRowByName(siglaEntita, info["SiglaInformazione"], Utility.Date.GetSuffissoDATA1), newNomiDefiniti.GetFirstCol() - 2).Extend(colOffset: 2 + Utility.Date.GetOreGiorno(Utility.DataBase.DataAttiva)));
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

        protected bool CreaDatiTopiciUnitaXML(object siglaEntita, object siglaAzione, string exportPath, DateTime dataRif)
        {
            try
            {
                string nomeFoglio = NewDefinedNames.GetSheetName(siglaEntita);
                NewDefinedNames newNomiDefiniti = new NewDefinedNames(nomeFoglio);
                Excel.Worksheet ws = Utility.Workbook.WB.Sheets[nomeFoglio];

                string suffissoData = Utility.Date.GetSuffissoData(dataRif);
                int oreGiorno = Utility.Date.GetOreGiorno(dataRif);

                DataView categoriaEntita = _localDB.Tables[Utility.DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";
                object codiceRUP = categoriaEntita[0]["CodiceRUP"];
                //bool isTermo = categoriaEntita[0]["SiglaCategoria"].Equals("IREN_60T");

                DataView entitaAzioneInformazione = _localDB.Tables[Utility.DataBase.Tab.ENTITA_AZIONE_INFORMAZIONE].DefaultView;
                entitaAzioneInformazione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";

                XElement datiTopici = new XElement("DatiTopiciUnit");

                XElement unit = new XElement("Unit", new XAttribute("StartDate", dataRif.ToString("yyyyMMdd")), new XAttribute("IDUnit", codiceRUP));

                for (int i = 0; i < oreGiorno; i++)
                {
                    XElement pr = new XElement("PR", i + 1);

                    foreach (DataRowView info in entitaAzioneInformazione)
                    {
                        object siglaEntitaRif = info["SiglaEntitaRif"] is DBNull ? siglaEntita : info["SiglaEntitaRif"];

                        Range rng = newNomiDefiniti.Get(siglaEntitaRif, info["SiglaInformazione"], suffissoData, Utility.Date.GetSuffissoOra(i + 1));
                        string value = (ws.Range[rng.ToString()].Value ?? "0").ToString().Replace('.', ',');

                        XAttribute attr = new XAttribute("CIAO", value);
                        pr.Add(attr);
                    }
                }

                XDocument datiTopiciUnita = new XDocument(new XDeclaration("1.0", "ISO-8859-1", "yes"),
                    new XElement("BMTransaction-DTU", new XAttribute("xmlns", "urn: XML-BIDMGM"), new XAttribute("xsi:schemaLocation", "urn:XML-BIDMGM BM_DatiTopiciUnita.xsd"), new XAttribute(XNamespace.Xmlns + "xsi", "http://www.w3.org/2001/XMLSchema-instance"), new XAttribute("ReferenceNumber", codiceRUP.ToString().Replace("_", "") + "_" + DateTime.Now.ToString("yyyyMMddHHmmss")), 
                        unit)
                    );

                string filename = "VDT_" + codiceRUP.ToString().ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xml";
                //variazioneDatiTecnici.Save(Path.Combine(exportPath, filename));

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

                DataView entitaProprieta = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'SISTEMA_COMANDI_ALLEGATO_EXCEL'";
                if (entitaProprieta.Count > 0)
                {
                    fileName = @"D:\" + entitaProprieta[0]["Valore"] + "_VDT_" + Utility.DataBase.DataAttiva.ToString("yyyyMMdd") + ".xls";

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

                    var config = Utility.Utilities.GetUsrConfigElement("destMailTest");
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

                    config = Utility.Utilities.GetUsrConfigElement("oggettoMail");
                    string oggetto = config.Value.Replace("%COD%", codUP).Replace("%DATA%", Utility.DataBase.DataAttiva.ToString("dd-MM-yyyy"));
                    config = Utility.Utilities.GetUsrConfigElement("messaggioMail");
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
                Utility.Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "SisCom - Esporta.InvioMail: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                if(File.Exists(fileName))
                    File.Delete(fileName);

                return false;
            }

            return true;
        }
    }
}
