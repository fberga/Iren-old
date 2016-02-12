using Iren.PSO;
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
                case "MAIL":
                    Workbook.ScreenUpdating = false;
                    DefinedNames mainDefinedNames = new DefinedNames("Main");
                    //TODO verificare se è sempre aggiornato
                    //unico caso che non aggiorna è se carico e faccio invia mail conseguentemente

                    Aggiorna a = new Aggiorna();
                    a.AggiornaPrevisioneRiepilogo();
                
                    //salvo i dati 
                    Riepilogo r = new Riepilogo();
                    r.SalvaPrevisione();

                    if (InviaMail(mainDefinedNames, siglaEntita))
                    {

                    }

                    Workbook.ScreenUpdating = true;
                    break;
            }
            return true;
        }

        protected bool InviaMail(DefinedNames definedNames, object siglaEntita) 
        {
            string fileName = "";
            try
            {
                fileName = @"D:\PrevisioneGAS_" + DateTime.Now.ToString("yyyyMMddmmss") + ".xls";

                Excel.Workbook wb = Globals.ThisWorkbook.Application.Workbooks.Add();

                Workbook.Main.Range[Range.GetRange(definedNames.GetFirstRow(), definedNames.GetFirstCol(), definedNames.GetRowOffset(), definedNames.GetColOffsetRiepilogo()).ToString()].Copy();
                wb.Sheets[1].Range["B2"].PasteSpecial();

                wb.Sheets[1].UsedRange.ColumnWidth = 17;
                wb.Sheets[1].Range["A1"].Select();
                wb.SaveAs(fileName, Excel.XlFileFormat.xlExcel8);
                wb.Close();
                Marshal.ReleaseComObject(wb);

                var config = Workbook.GetUsrConfigElement("destMailTest");
                string mailTo = config.Value;
                string mailCC = "";

                DataView entitaProprieta = new DataView(Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA]);

                if (Workbook.Ambiente == Simboli.PROD)
                {
                    entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'PREV_CONSUMO_GAS_MAIL_TO' AND IdApplicazione = " + Workbook.IdApplicazione;
                    mailTo = entitaProprieta[0]["Valore"].ToString();
                    entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'PREV_CONSUMO_GAS_MAIL_CC' AND IdApplicazione = " + Workbook.IdApplicazione;
                    mailCC = entitaProprieta[0]["Valore"].ToString();
                }

                Outlook.Application outlook = GetOutlookInstance();
                Outlook._MailItem mail = outlook.CreateItem(Outlook.OlItemType.olMailItem);

                config = Workbook.GetUsrConfigElement("oggettoMail");
                string oggetto = config.Value.Replace("%DATA%", Workbook.DataAttiva.ToString("dd-MM-yyyy"));
                config = Workbook.GetUsrConfigElement("messaggioMail");
                string messaggio = config.Value;
                messaggio = Regex.Replace(messaggio, @"^[^\S\r\n]+", "", RegexOptions.Multiline);


                ////TODO check se manda sempre con lo stesso account...
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
            catch(Exception e)
            {
                Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "PrevisioneGAS.Esporta.InvioMail: " + e.Message);

                System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                if(File.Exists(fileName))
                    File.Delete(fileName);

                return false;
            }

            return true;
        }
    }
}
