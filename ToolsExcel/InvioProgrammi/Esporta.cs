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
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Iren.ToolsExcel
{
    class Esporta : AEsporta
    {
        DefinedNames _defNamesMercato = new DefinedNames(Simboli.Mercato);

        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            DataView entitaAzione = _localDB.Tables[Utility.DataBase.Tab.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "'";
            if (entitaAzione.Count == 0)
                return false;

            switch (siglaAzione.ToString())
            {
                case "MAIL":
                    //carico i path di export
                    Dictionary<UserConfigElement, string> cfgPaths = new Dictionary<UserConfigElement, string>();
                    var path = Utility.Workbook.GetUsrConfigElement("pathExportFileFMS");
                    cfgPaths.Add(path, Utility.ExportPath.PreparePath(path.Value));
                    path = Utility.Workbook.GetUsrConfigElement("pathExportFileXSD");
                    cfgPaths.Add(path, Utility.ExportPath.PreparePath(path.Value));
                    path = Utility.Workbook.GetUsrConfigElement("pathExportFileRS");
                    cfgPaths.Add(path, Utility.ExportPath.PreparePath(path.Value));

                    //verifico che siano tutti raggiungibili
                    foreach (var kv in cfgPaths)
                    {
                        if(!Directory.Exists(kv.Value))
                        {
                            System.Windows.Forms.MessageBox.Show(path.Desc + " '" + kv.Value + "' non raggiungibile.", Simboli.nomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                            return false;
                        }
                    }

                    Globals.ThisWorkbook.Application.ScreenUpdating = false;

                    var oldActiveWindow = Globals.ThisWorkbook.Application.ActiveWindow;
                    Globals.ThisWorkbook.Worksheets[Simboli.Mercato].Activate();

                    Range rng = new Range(_defNamesMercato.GetRowByName(siglaEntita, "DATA"), 1, Utility.Date.GetOreGiorno(Utility.DataBase.DataAttiva) + 5, _defNamesMercato.GetLastCol());

                    InviaMail(Simboli.Mercato, siglaEntita, rng);
                 
                    oldActiveWindow.Activate();

                    Globals.ThisWorkbook.Application.ScreenUpdating = true;
                    break;
            }
            return true;
        }
        protected bool CreaOutputXLS(Excel.Worksheet ws, string fileName, bool deleteOrco, Range rng)
        {
            bool hasVariations = false;

            Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Add();
            ws.Range[rng.ToString()].Copy();
            wb.Sheets[1].Range["A1"].PasteSpecial(Excel.XlPasteType.xlPasteAllUsingSourceTheme);

            //fisso la formattazione condizionale nel range copiato
            foreach (Range cell in rng.Cells)
            {
                //traslo la cella per il nuovo foglio
                Range tCell = new Range(cell);
                tCell.StartRow -= (rng.StartRow - 1);
                tCell.StartColumn -= (rng.StartColumn - 1);
                wb.Sheets[1].Range[tCell.ToString()].Interior.ColorIndex = ws.Range[cell.ToString()].DisplayFormat.Interior.ColorIndex;

                if (wb.Sheets[1].Range[tCell.ToString()].Interior.ColorIndex == Struct.COLORE_VARIAZIONE_NEGATIVA || wb.Sheets[1].Range[tCell.ToString()].Interior.ColorIndex == Struct.COLORE_VARIAZIONE_POSITIVA)
                    hasVariations = true;
            }
            //rimuovo la formattazione condizionale
            Excel.Range tab = wb.Sheets[1].UsedRange;
            tab.FormatConditions.Delete();

            if (deleteOrco)
            {
                //TODO CHECK se rimuovere...
                if(DateTime.Now < new DateTime(2014, 07, 01))
                    wb.Sheets[1].Columns[8].EntireColumn.Delete();

                wb.Sheets[1].Range["H3"].Value = "Programma indicativo ORCO";
                wb.Sheets[1].Range["H3"].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            }

            //salvo l'export e lo chiudo
            wb.Sheets[1].Range["A1"].Select();
            wb.SaveAs(fileName, Excel.XlFileFormat.xlExcel8);
            wb.Close();

            return hasVariations;
        }
        protected bool InviaMail(string nomeFoglio, object siglaEntita, Range rng) 
        {
            List<string> attachments = new List<string>();
            bool hasVariations = false;
            try
            {
                Excel.Worksheet ws = Globals.ThisWorkbook.Sheets[nomeFoglio];
                
                //inizializzo l'oggetto mail
                Outlook.Application outlook = GetOutlookInstance();
                Outlook.MailItem mail = outlook.CreateItem(Outlook.OlItemType.olMailItem);

                DataView entitaProprieta = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_EXCEL'";
                if (entitaProprieta.Count > 0)
                {
                    //creo file Excel da allegare
                    attachments.Add(@"D:\" + Utility.DataBase.DataAttiva.ToString("yyyyMMdd") + "_" + entitaProprieta[0]["Valore"] + "_" + Simboli.Mercato + ".xls");

                    hasVariations = CreaOutputXLS(ws, attachments.Last(), siglaEntita.Equals("CE_ORX"), rng);


                    DataView categoriaEntita = Utility.DataBase.LocalDB.Tables[Utility.DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                    categoriaEntita.RowFilter = "Gerarchia = '" + siglaEntita + "'";

                    if(categoriaEntita.Count == 0)
                        categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "'";

                    foreach (DataRowView entita in categoriaEntita)
                    {
                        entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_FMS'";
                        if (entitaProprieta.Count > 0)
                        {
                            //cerco i file XML
                            string nomeFileFMS = Utility.ExportPath.PreparePath(Utility.Workbook.GetUsrConfigElement("formatoNomeFileFMS").Value, codRup: entita["CodiceRup"].ToString()) + "*.xml";
                            string pathFileFMS = Utility.Workbook.GetUsrConfigElement("pathExportFileFMS").Value;

                            string[] files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);

                            foreach (string file in files)
                                attachments.Add(file);
                        }

                        entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_FMS'";
                        if (entitaProprieta.Count > 0)
                        {
                            //cerco i file XML
                            string nomeFileFMS = Utility.ExportPath.PreparePath(Utility.Workbook.GetUsrConfigElement("formatoNomeFileFMS").Value, codRup: entita["CodiceRup"].ToString()) + "*.xml";
                            string pathFileFMS = Utility.Workbook.GetUsrConfigElement("pathExportFileFMS").Value;

                            string[] files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);

                            if (files.Length > 0)
                            {
                                foreach (string file in files)
                                    attachments.Add(file);
                            }
                            else
                            {
                                nomeFileFMS = Utility.ExportPath.PreparePath(Utility.Workbook.GetUsrConfigElement("formatoNomeFileFMS_TERNA").Value, codRup: entita["CodiceRup"].ToString()) + "*.xml";
                                pathFileFMS = Utility.Workbook.GetUsrConfigElement("pathExportFileFMS").Value;

                                files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);
                                foreach (string file in files)
                                    attachments.Add(file);
                            }
                        }
                        entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_RS'";
                        if (entitaProprieta.Count > 0)
                        {
                            string nomeFileFMS = Utility.ExportPath.PreparePath(Utility.Workbook.GetUsrConfigElement("formatoNomeFileRS_TERNA").Value) + ".xml";
                            string pathFileFMS = Utility.Workbook.GetUsrConfigElement("pathExportFileRS").Value;

                            string[] files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);
                            foreach (string file in files)
                                attachments.Add(file);
                        }
                    }
                    


                    var config = Utility.Workbook.GetUsrConfigElement("destMailTest");
                    string mailTo = config.Value;
                    string mailCC = "";

                    if (Simboli.Ambiente == "Produzione")
                    {
                        entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_MAIL_TO'";
                        mailTo = entitaProprieta[0]["Valore"].ToString();
                        entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_MAIL_CC'";
                        mailCC = entitaProprieta[0]["Valore"].ToString();
                    }

                    entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_CODICE_MAIL'";
                    string codUP = entitaProprieta[0]["Valore"].ToString();

                    config = Utility.Workbook.GetUsrConfigElement("oggettoMail");
                    string oggetto = config.Value.Replace("%COD%", codUP).Replace("%DATA%", Utility.DataBase.DataAttiva.ToString("dd-MM-yyyy")).Replace("%MSD%", Simboli.Mercato) + (hasVariations ? " - CON VARIAZIONI" : "");
                    config = Utility.Workbook.GetUsrConfigElement("messaggioMail");
                    string messaggio = config.Value;
                    messaggio = Regex.Replace(messaggio, @"^[^\S\r\n]+", "", RegexOptions.Multiline);

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

                    //aggiungo allegato XLS
                    foreach (string attachment in attachments)
                        mail.Attachments.Add(attachment);

                    mail.Send();

                    foreach (string file in attachments)
                        File.Delete(file);
                }
            }
            catch(Exception e)
            {
                Utility.Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "InvioProgrammi - Esporta.InvioMail: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                foreach (string file in attachments)
                    File.Delete(file);

                return false;
            }

            return true;
        }
    }
}
