using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Data;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Iren.FrontOffice.Core;
using System.Text.RegularExpressions;
using System.Configuration;
using Word = Microsoft.Office.Interop.Word;
using DataTable = System.Data.DataTable;
using DataRow = System.Data.DataRow;
using DataView = System.Data.DataView;
using RiMoST2.Properties;
using System.Drawing;

namespace RiMoST2
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        #region Variabili

        private Office.IRibbonUI ribbon;
        FormAnnullaModifica _formAnnullaModifica;

        #endregion

        #region Costruttori

        public Ribbon()
        {
        }

        #endregion

        #region Metodi Privati

        private void getAvailableID()
        {
            DataTable dt = ThisDocument._db.Select("spGetFirstAvailableID");
            Globals.ThisDocument.lbIdRichiesta.Text = dt.Rows[0][0].ToString();
        }

        private void Highlight(string textToFind, Word.WdColorIndex color, string highlightMark = "")
        {
            Word.Find finder = Globals.ThisDocument.Content.Find;
            finder.Text = textToFind;
            finder.Replacement.Text = finder.Text + highlightMark;
            finder.Replacement.Font.ColorIndex = color;
            finder.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }
        private void ToNormal(string textToFind, Word.WdColorIndex color, string highlightMark = "")
        {
            Word.Find finder = Globals.ThisDocument.Content.Find;
            finder.Text = textToFind + highlightMark;
            finder.Replacement.Text = textToFind;
            finder.Replacement.Font.ColorIndex = color;
            finder.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }

        private void Print()
        {
            object missing = Missing.Value;

            if (Globals.ThisDocument.Application.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show() == 1)
            {
                Globals.ThisDocument.PrintOut(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
        }

        #endregion

        #region Membri IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("RiMoST2.Ribbon.xml");
        }

        #endregion

        #region Callback della barra multifunzione

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public void btnReset_Click(Office.IRibbonControl control)
        {
            if (MessageBox.Show("Sicuro di voler cancellare il contenuto dei campi?", "Cancellare campi?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                Globals.ThisDocument.cmbStrumento.SelectedIndex = 0;
                Globals.ThisDocument.txtDescrizione.Text = "";
                Globals.ThisDocument.txtOggetto.Text = "";
                Globals.ThisDocument.txtNote.Text = "";
                Globals.ThisDocument.dtDataCreazione.Value = DateTime.Now;

                getAvailableID();
            }
        }
        public void btnInvia_Click(Office.IRibbonControl control)
        {
            object copies = "1";
            object pages = "";
            object range = Word.WdPrintOutRange.wdPrintAllDocument;
            object items = Word.WdPrintOutItem.wdPrintDocumentContent;
            object pageType = Word.WdPrintOutPages.wdPrintAllPages;
            object oTrue = true;
            object oFalse = false;
            object missing = Missing.Value;

            QryParams parameters = new QryParams()
            {
                {"@IdRichiesta", Globals.ThisDocument.lbIdRichiesta.Text}
            };

            DataTable dt = ThisDocument._db.Select("spGetRichiesta", parameters);
            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("Esiste già una richiesta con lo stesso codice. Premere sul tasto di refresh per ottenerne uno nuovo", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (Globals.ThisDocument.txtOggetto.Text == "" || Globals.ThisDocument.txtDescrizione.Text == "")
                {
                    MessageBox.Show("Alcuni campi obbligatori non sono stati compilati. Compilare i campi evidenziati!", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    Globals.ThisDocument.RemoveProtection();
                    Globals.ThisDocument.Application.ScreenUpdating = false;

                    ToNormal("Oggetto", Word.WdColorIndex.wdBlack, "*");
                    ToNormal("Descrizione", Word.WdColorIndex.wdBlack, "*");

                    if (Globals.ThisDocument.txtOggetto.Text == "")
                        Highlight("Oggetto", Word.WdColorIndex.wdRed, "*");

                    if (Globals.ThisDocument.txtDescrizione.Text == "")
                        Highlight("Descrizione", Word.WdColorIndex.wdRed, "*");

                    Globals.ThisDocument.Application.ScreenUpdating = true;
                    Globals.ThisDocument.AddProtection();
                }
                else
                {
                    Globals.ThisDocument.RemoveProtection();
                    Globals.ThisDocument.Application.ScreenUpdating = false;

                    ToNormal("Oggetto", Word.WdColorIndex.wdBlack, "*");
                    ToNormal("Descrizione", Word.WdColorIndex.wdBlack, "*");

                    Globals.ThisDocument.Application.ScreenUpdating = true;
                    Globals.ThisDocument.AddProtection();

                    if (MessageBox.Show("Sicuro di voler inviare il documento?", "Stampa e invia?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                    {
                        Regex rgx = new Regex(@"(\[[^\[\]]*\])");
                        string saveName = ConfigurationManager.AppSettings["saveNameFormat"];

                        foreach (Match m in rgx.Matches(saveName))
                        {
                            try
                            {
                                Control c = (Control)Globals.ThisDocument.Controls[m.Value.Replace("[", "").Replace("]", "")];
                                saveName = saveName.Replace(m.Value, c.Text);
                            }
                            catch (ArgumentOutOfRangeException)
                            {

                            }
                        }
                        rgx = new Regex(@"([^\.\-_a-zA-Z0-9]+)");

                        string name = rgx.Replace(saveName, "_");

                        object savePath = Path.Combine(ConfigurationManager.AppSettings["savePath"], name + ".pdf");
                        object format = Word.WdSaveFormat.wdFormatPDF;
                        try
                        {
                            Globals.ThisDocument.SaveAs2(ref savePath, ref format, ref oTrue, ref missing, ref oFalse,
                                ref missing, ref oFalse, ref missing, ref missing, ref oFalse, ref oFalse, ref missing,
                                ref missing, ref missing, ref missing, ref missing, ref missing);

                            DateTime dataInvio = DateTime.Parse(Globals.ThisDocument.lbDataInvio.Text);
                            DataRowView strumento = (DataRowView)Globals.ThisDocument.cmbStrumento.SelectedItem;

                            parameters = new QryParams()
                            {
                                {"@IdRichiesta", Globals.ThisDocument.lbIdRichiesta.Text},
                                {"@DataCreazione", Globals.ThisDocument.dtDataCreazione.Value.ToString("yyyyMMdd")},
                                {"@DataInvio", dataInvio.ToString("yyyyMMdd")},
                                {"@IdApplicazione", strumento["IdApplicazione"]},
                                {"@Oggetto", Globals.ThisDocument.txtOggetto.Text},
                                {"@Descr", Globals.ThisDocument.txtDescrizione.Text},
                                {"@Note", Globals.ThisDocument.txtNote.Text},
                                {"@NomeFile", savePath}
                            };


                            ThisDocument._db.Insert("spAddNewRichiestaModifica", parameters);

                            Print();
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Salvataggio non riuscito... Riprovare più tardi.", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }
        public void btnChiudi_Click(Office.IRibbonControl control)
        {
            Globals.ThisDocument.CloseWithoutSaving();
        }
        public void btnRefresh_Click(Office.IRibbonControl control)
        {
            getAvailableID();
        }
        public void btnPrint_Click(Office.IRibbonControl control)
        {
            Print();
        }
        public void btnAnnulla_Click(Office.IRibbonControl control)
        {
            if (_formAnnullaModifica == null || _formAnnullaModifica.IsDisposed)
            {
                _formAnnullaModifica = new FormAnnullaModifica();
                _formAnnullaModifica.Show();
            }
            _formAnnullaModifica.WindowState = FormWindowState.Normal;
            _formAnnullaModifica.Focus();
        }

        public Bitmap btnReset_getImage(Office.IRibbonControl control)
        {
            return Resources.Eraser_icon;
        }
        public Bitmap btnInvia_getImage(Office.IRibbonControl control)
        {
            return Resources.Send_icon;
        }
        public Bitmap btnChiudi_getImage(Office.IRibbonControl control)
        {
            return Resources.Close_icon;
        }
        public Bitmap btnRefresh_getImage(Office.IRibbonControl control)
        {
            return Resources.Refresh_icon;
        }
        public Bitmap btnPrint_getImage(Office.IRibbonControl control)
        {
            return Resources.Print_icon;
        }
        public Bitmap btnAnnulla_getImage(Office.IRibbonControl control)
        {
            return Resources.Bin_icon;
        }

        #endregion

        #region Supporti

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
