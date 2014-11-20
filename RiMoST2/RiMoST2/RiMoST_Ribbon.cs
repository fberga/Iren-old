using System;
using System.IO;
using System.Configuration;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Reflection;
using System.Text.RegularExpressions;
using DataTable = System.Data.DataTable;
using DataRow = System.Data.DataRow;
using DataView = System.Data.DataView;
using DataRowView = System.Data.DataRowView;
using Iren.FrontOffice.Core;
using System.Deployment.Application;

namespace RiMoST2
{
    public partial class RiMoST_Ribbon
    {
        FormAnnullaModifica _formAnnullaModifica;

        private System.Version getCurrentV()
        {
            try
            {
                return ApplicationDeployment.CurrentDeployment.CurrentVersion;
            }
            catch (Exception)
            {
                return Assembly.GetExecutingAssembly().GetName().Version;
            }
        }

        private void RiMoST_Load(object sender, RibbonUIEventArgs e)
        {
            DataTable dt = ThisDocument._db.Select("spGetAvailableYears");
            foreach (DataRow r in dt.Rows) 
            {
                RibbonDropDownItem i = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();;
                i.Label = r["Anno"].ToString();
                cbAnniDisponibili.Items.Add(i);
            }
            cbAnniDisponibili.Text = cbAnniDisponibili.Items[0].Label;

            System.Version appV = getCurrentV();
            lbVersioneApp.Label = "  App v" + appV.ToString();

            System.Version CoreV = ThisDocument._db.GetCurrentV();
            lbCoreV.Label = "  Core v" + CoreV.ToString();

            
            
        }

        private void getAvailableID()
        {
            DataTable dt = ThisDocument._db.Select("spGetFirstAvailableID");
            Globals.ThisDocument.lbIdRichiesta.Text = dt.Rows[0][0].ToString();
        }

        private void btnReset_Click(object sender, RibbonControlEventArgs e)
        {
            if (MessageBox.Show("Sicuro di voler cancellare il contenuto dei campi?", "Cancellare campi?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                Globals.ThisDocument.cmbStrumento.SelectedIndex = 0;
                ((DataView)Globals.ThisDocument.cmbStrumento.DataSource).RowFilter = "";
                chkIsDraft.Checked = false;
                Globals.ThisDocument.cmbStrumento.Enabled = true;
                Globals.ThisDocument.txtDescrizione.Text = "";
                Globals.ThisDocument.txtOggetto.Text = "";
                Globals.ThisDocument.txtNote.Text = "";
                Globals.ThisDocument.dtDataCreazione.Value = DateTime.Now;
                btnRefresh.Enabled = true;
                btnSalvaBozza.Enabled = true;
                chkIsDraft.Checked = false;

                getAvailableID();
            }
        }

        private void Highlight(string textToFind, WdColorIndex color, string highlightMark = "")
        {
            Find finder = Globals.ThisDocument.Content.Find;
            finder.Text = textToFind;
            finder.Replacement.Text = finder.Text + highlightMark;
            finder.Replacement.Font.ColorIndex = color;
            finder.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }
        private void ToNormal(string textToFind, WdColorIndex color, string highlightMark = "")
        {
            Find finder = Globals.ThisDocument.Content.Find;
            finder.Text = textToFind + highlightMark;
            finder.Replacement.Text = textToFind;
            finder.Replacement.Font.ColorIndex = color;
            finder.Execute(Replace: Word.WdReplace.wdReplaceAll);
        }

        private void btnInvia_Click(object sender, RibbonControlEventArgs e)
        {
            if (chkIsDraft.Checked)
            {
                MessageBox.Show("La richiesta è contrassegnata come bozza. Togliere la spunta e riprovare.", "Impossibile salvare!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
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

                DataView dv = ThisDocument._db.Select("spGetRichiesta", parameters).DefaultView;
                dv.RowFilter = "IdTipologiaStato <> 7";
                if (dv.Count > 0)
                {
                    MessageBox.Show("Esiste già una richiesta con lo stesso codice. Premere sul tasto di refresh per ottenerne uno nuovo", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (!EmptyFields())
                    {
                        if (MessageBox.Show("Sicuro di voler inviare il documento?", "Stampa e invia?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                        {
                            btnSalvaBozza.Enabled = false;

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
                            object format = WdSaveFormat.wdFormatPDF;
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
                                {"@IdTipologiaStato", 1},
                                {"@IdApplicazione", strumento["IdApplicazione"]},
                                {"@Oggetto", Globals.ThisDocument.txtOggetto.Text},
                                {"@Descr", Globals.ThisDocument.txtDescrizione.Text},
                                {"@Note", Globals.ThisDocument.txtNote.Text},
                                {"@NomeFile", savePath}
                            };


                                ThisDocument._db.Insert("spSaveRichiestaModifica", parameters);

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
        }

        private void btnChiudi_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisDocument.CloseWithoutSaving();
        }

        private void btnRefresh_Click(object sender, RibbonControlEventArgs e)
        {
            getAvailableID();
        }

        private void btnPrint_Click(object sender, RibbonControlEventArgs e)
        {
            Print();
        }

        private void Print()
        {
            object missing = Missing.Value;

            if (Globals.ThisDocument.Application.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show() == 1)
            {
                Globals.ThisDocument.PrintOut(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
        }

        private void btnAnnulla_Click(object sender, RibbonControlEventArgs e)
        {
            if (_formAnnullaModifica == null || _formAnnullaModifica.IsDisposed)
            {
                _formAnnullaModifica = new FormAnnullaModifica("");
                _formAnnullaModifica.Show();
            }
            _formAnnullaModifica.WindowState = FormWindowState.Normal;
            _formAnnullaModifica.Focus();
        }

        private bool EmptyFields()
        {
            if (Globals.ThisDocument.txtOggetto.Text == "" || Globals.ThisDocument.txtDescrizione.Text == "")
            {
                MessageBox.Show("Alcuni campi obbligatori non sono stati compilati. Compilare i campi evidenziati!", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Error);

                Globals.ThisDocument.RemoveProtection();
                Globals.ThisDocument.Application.ScreenUpdating = false;

                ToNormal("Oggetto", WdColorIndex.wdBlack, "*");
                ToNormal("Descrizione", WdColorIndex.wdBlack, "*");

                if (Globals.ThisDocument.txtOggetto.Text == "")
                    Highlight("Oggetto", WdColorIndex.wdRed, "*");

                if (Globals.ThisDocument.txtDescrizione.Text == "")
                    Highlight("Descrizione", WdColorIndex.wdRed, "*");

                Globals.ThisDocument.Application.ScreenUpdating = true;
                Globals.ThisDocument.AddProtection();

                return true;
            }

            Globals.ThisDocument.RemoveProtection();
            Globals.ThisDocument.Application.ScreenUpdating = false;

            ToNormal("Oggetto", WdColorIndex.wdBlack, "*");
            ToNormal("Descrizione", WdColorIndex.wdBlack, "*");

            Globals.ThisDocument.Application.ScreenUpdating = true;
            Globals.ThisDocument.AddProtection();
            
            return false;
        }

        private void btnSalvaBozza_Click(object sender, RibbonControlEventArgs e)
        {
            if (!EmptyFields())
            {
                DateTime dataInvio = DateTime.Parse(Globals.ThisDocument.lbDataInvio.Text);
                DataRowView strumento = (DataRowView)Globals.ThisDocument.cmbStrumento.SelectedItem;

                chkIsDraft.Checked = true;

                QryParams parameters = new QryParams()
                {
                    {"@IdRichiesta", Globals.ThisDocument.lbIdRichiesta.Text},
                    {"@DataCreazione", Globals.ThisDocument.dtDataCreazione.Value.ToString("yyyyMMdd")},
                    {"@DataInvio", dataInvio.ToString("yyyyMMdd")},
                    {"@IdTipologiaStato", 7},
                    {"@IdApplicazione", strumento["IdApplicazione"]},
                    {"@Oggetto", Globals.ThisDocument.txtOggetto.Text},
                    {"@Descr", Globals.ThisDocument.txtDescrizione.Text},
                    {"@Note", Globals.ThisDocument.txtNote.Text}
                };

                try
                {
                    ThisDocument._db.Insert("spSaveRichiestaModifica", parameters);
                }
                catch (Exception)
                {
                    MessageBox.Show("Salvataggio non riuscito... Riprovare più tardi.", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnModifica_Click(object sender, RibbonControlEventArgs e)
        {
            //SelezionaModifica selMod = new SelezionaModifica("");
            //selMod.ShowDialog();
        }
    }
}
