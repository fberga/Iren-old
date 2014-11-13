using System;
using System.IO;
using System.Configuration;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using DataTable = System.Data.DataTable;
using DataRow = System.Data.DataRow;
using DataView = System.Data.DataView;

namespace RiMoST2
{
    public partial class RiMoST
    {
        private void RiMoST_Load(object sender, RibbonUIEventArgs e)
        {
            //IList<IRibbonExtension> ribbonList = Globals.Ribbons.Base;

            //foreach (IRibbonExtension r in ribbonList)
            //{
            //    r.ToString();
            //}
        }

        private void btnReset_Click(object sender, RibbonControlEventArgs e)
        {
            if (MessageBox.Show("Sicuro di voler cancellare il contenuto dei campi?", "Cancellare campi?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                Globals.ThisDocument.cmbStrumento.SelectedIndex = 0;
                Globals.ThisDocument.txtDescrizione.Text = "";
                Globals.ThisDocument.txtOggetto.Text = "";
                Globals.ThisDocument.txtNote.Text = "";
                Globals.ThisDocument.dtDataCreazione.Value = DateTime.Now;
                
                DataTable dt = ThisDocument._db.Select("spGetFirstAvailableID");
                Globals.ThisDocument.lbIdRichiesta.Text = dt.Rows[0][0].ToString();
            }
        }

        private void btnInvia_Click(object sender, RibbonControlEventArgs e)
        {
            Dictionary<string, object> parameters = new Dictionary<string, object>()
                {
                    {"@IdRichiesta", Globals.ThisDocument.lbIdRichiesta.Text}
                };

            DataTable dt = ThisDocument._db.Select("spGetRichiesta");
            if (dt.Rows.Count > 0) 
            {
                MessageBox.Show("Esiste già una richiesta con lo stesso codice. Premere sul tasto di refresh per ottenerne uno nuovo", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (MessageBox.Show("Sicuro di voler inviare il documento?", "Stampa e invia?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    object copies = "1";
                    object pages = "";
                    object range = Word.WdPrintOutRange.wdPrintAllDocument;
                    object items = Word.WdPrintOutItem.wdPrintDocumentContent;
                    object pageType = Word.WdPrintOutPages.wdPrintAllPages;
                    object oTrue = true;
                    object oFalse = false;
                    object missing = Missing.Value;

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

                        /*DateTime dataInvio = DateTime.Parse(Globals.ThisDocument.lbDataInvio.Text);
                        DataRowView strumento = (DataRowView)Globals.ThisDocument.cmbStrumento.SelectedItem;

                        Dictionary<string, object> parameters = new Dictionary<string,object>()
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


                        ThisDocument._db.Insert("spAddNewRichiestaModifica", parameters);*/

                        if (Globals.ThisDocument.Application.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show() == 1)
                        {
                            Globals.ThisDocument.PrintOut(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                        }
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Salvataggio non riuscito... Riprovare più tardi.", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnChiudi_Click(object sender, RibbonControlEventArgs e)
        {
            object saveMod = WdSaveOptions.wdDoNotSaveChanges;
            object missing = Missing.Value;
            Globals.ThisDocument.Close(ref saveMod, ref missing, ref missing);
        }
    }
}
