using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.UserConfig;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Iren.ToolsExcel
{
    /// <summary>
    /// Crea la mail con i dati di export da inviare agli impianti.
    /// </summary>
    class Esporta : AEsporta
    {
        DefinedNames _defNamesMercato = new DefinedNames(Simboli.Mercato);

        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            DataView entitaAzione = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE].DefaultView;
            entitaAzione.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaAzione = '" + siglaAzione + "' AND IdApplicazione = " + Simboli.AppID;
            if (entitaAzione.Count == 0)
                return false;

            switch (siglaAzione.ToString())
            {
                case "MAIL":
                    //carico i path di export
                    List<UserConfigElement> cfgPaths = new List<UserConfigElement>();

                    var cfgPath = Workbook.GetUsrConfigElement("pathExportFileFMS");
                    cfgPaths.Add(cfgPath);
                    cfgPath = Workbook.GetUsrConfigElement("pathExportFileXSD");
                    cfgPaths.Add(cfgPath);
                    cfgPath = Workbook.GetUsrConfigElement("pathExportFileRS");
                    cfgPaths.Add(cfgPath);

                    //verifico che siano tutti raggiungibili
                    foreach (var p in cfgPaths)
                    {
                        string path = PreparePath(p);
                        if (!Directory.Exists(path))
                        {
                            System.Windows.Forms.MessageBox.Show(p.Desc + " '" + path + "' non raggiungibile.", Simboli.nomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                            return false;
                        }
                    }

                    Globals.ThisWorkbook.Application.ScreenUpdating = false;

                    var oldActiveWindow = Globals.ThisWorkbook.Application.ActiveWindow;
                    Globals.ThisWorkbook.Worksheets[Simboli.Mercato].Activate();

                    Range rng = new Range(_defNamesMercato.GetRowByName(siglaEntita, "DATA"), 1, Date.GetOreGiorno(DataBase.DataAttiva) + 5, _defNamesMercato.GetLastCol());

                    bool result = InviaMail(Simboli.Mercato, siglaEntita, rng);
                 
                    oldActiveWindow.Activate();

                    Globals.ThisWorkbook.Application.ScreenUpdating = true;
                    return result;
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

                DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
                entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_EXCEL' AND IdApplicazione = " + Simboli.AppID;
                if (entitaProprieta.Count > 0)
                {
                    //creo file Excel da allegare
                    string pathExport = PreparePath(Workbook.GetUsrConfigElement("exportXML"));
                    attachments.Add(Path.Combine(pathExport, DataBase.DataAttiva.ToString("yyyyMMdd") + "_" + entitaProprieta[0]["Valore"] + "_" + Simboli.Mercato + ".xls"));

                    hasVariations = CreaOutputXLS(ws, attachments.Last(), siglaEntita.Equals("CE_ORX"), rng);


                    DataView categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
                    categoriaEntita.RowFilter = "Gerarchia = '" + siglaEntita + "' AND IdApplicazione = " + Simboli.AppID;

                    if(categoriaEntita.Count == 0)
                        categoriaEntita.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND IdApplicazione = " + Simboli.AppID;

                    bool interrupt = false;

                    foreach (DataRowView entita in categoriaEntita)
                    {
                        entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_FMS' AND IdApplicazione = " + Simboli.AppID;
                        if (entitaProprieta.Count > 0)
                        {
                            //cerco i file XML
                            string nomeFileFMS = PrepareName(Workbook.GetUsrConfigElement("formatoNomeFileFMS").Value, codRup: entita["CodiceRup"].ToString()) + "*.xml";
                            string pathFileFMS = PreparePath(Workbook.GetUsrConfigElement("pathExportFileFMS"));

                            string[] files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);

                            if (files.Length == 0)
                            {
                                if (System.Windows.Forms.MessageBox.Show("File FMS non presente nell'area di rete. Continuare con l'invio?", Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                                {
                                    interrupt = true;
                                    break;
                                }
                            }
                            foreach (string file in files)
                                attachments.Add(file);
                        }

                        entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_FMS' AND IdApplicazione = " + Simboli.AppID;
                        if (entitaProprieta.Count > 0)
                        {
                            //cerco i file XML
                            string nomeFileFMS = PrepareName(Workbook.GetUsrConfigElement("formatoNomeFileFMS").Value, codRup: entita["CodiceRup"].ToString()) + "*.xml";
                            string pathFileFMS = PreparePath(Workbook.GetUsrConfigElement("pathExportFileFMS"));

                            string[] files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);



                            if (files.Length > 0)
                            {
                                foreach (string file in files)
                                    attachments.Add(file);
                            }
                            else
                            {
                                nomeFileFMS = PrepareName(Workbook.GetUsrConfigElement("formatoNomeFileFMS_TERNA").Value, codRup: entita["CodiceRup"].ToString()) + "*.xml";
                                pathFileFMS = PreparePath(Workbook.GetUsrConfigElement("pathExportFileFMS"));

                                files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);

                                if (files.Length == 0)
                                {
                                    if (System.Windows.Forms.MessageBox.Show("File FMS non presente nell'area di rete. Continuare con l'invio?", Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                                    {
                                        interrupt = true;
                                        break;
                                    }
                                }

                                foreach (string file in files)
                                    attachments.Add(file);
                            }
                        }
                        entitaProprieta.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_ALLEGATO_RS' AND IdApplicazione = " + Simboli.AppID;
                        if (entitaProprieta.Count > 0)
                        {
                            string nomeFileFMS = PrepareName(Workbook.GetUsrConfigElement("formatoNomeFileRS_TERNA").Value) + ".xml";
                            string pathFileFMS = PreparePath(Workbook.GetUsrConfigElement("pathExportFileRS"));

                            string[] files = Directory.GetFiles(pathFileFMS, nomeFileFMS, SearchOption.TopDirectoryOnly);

                            if (files.Length == 0)
                            {
                                if (System.Windows.Forms.MessageBox.Show("File Riserva Secondaria non presente nell'area di rete. Continuare con l'invio?", Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                                {
                                    interrupt = true;
                                    break;
                                }
                            }

                            foreach (string file in files)
                                attachments.Add(file);
                        }
                    }

                    if (!interrupt)
                    {
                        var config = Workbook.GetUsrConfigElement("destMailTest");
                        string mailTo = config.Value;
                        string mailCC = "";

                        if (Simboli.Ambiente == "Produzione")
                        {
                            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_MAIL_TO' AND IdApplicazione = " + Simboli.AppID;
                            mailTo = entitaProprieta[0]["Valore"].ToString();
                            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_MAIL_CC' AND IdApplicazione = " + Simboli.AppID;
                            mailCC = entitaProprieta[0]["Valore"].ToString();
                        }

                        entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'INVIO_PROGRAMMA_CODICE_MAIL' AND IdApplicazione = " + Simboli.AppID;
                        string codUP = entitaProprieta[0]["Valore"].ToString();

                        config = Workbook.GetUsrConfigElement("oggettoMail");
                        string oggetto = config.Value.Replace("%COD%", codUP).Replace("%DATA%", DataBase.DataAttiva.ToString("dd-MM-yyyy")).Replace("%MSD%", Simboli.Mercato) + (hasVariations ? " - CON VARIAZIONI" : "");
                        config = Workbook.GetUsrConfigElement("messaggioMail");
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
                    }
                    
                    foreach (string file in attachments)
                        File.Delete(file);

                    return !interrupt;
                }
            }
            catch(Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "InvioProgrammi - Esporta.InvioMail: " + e.Message);
                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.nomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                foreach (string file in attachments)
                    File.Delete(file);

                return false;
            }

            return false;
        }
    }
}
