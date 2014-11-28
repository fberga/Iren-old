﻿using System;
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
using System.Drawing;
using System.Deployment.Application;

namespace Iren.FrontOffice.Tools
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        #region Variabili

        private Office.IRibbonUI ribbon;
        FormAnnullaModifica _formAnnullaModifica;
        internal int _cbAnniDispCount = 0;
        internal int _cbAnniDispIndex = 0;
        internal List<string> _cbAnniDispLabels;
        internal string _cbAnniDispValue = "";
        internal System.Version _appV;
        internal System.Version _coreV;
        internal bool _chkIsDraftEnabled = true;
        internal bool _chkIsDraft = false;
        internal bool _btnSalvaBozzaEnabled = true;
        internal bool _btnRefreshEnabled = true;
        public static DataBase _db;

        #endregion

        #region Costruttori

        public Ribbon(ref DataBase db)
        {
            _db = db;
        }

        #endregion

        #region Metodi Privati

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

        private bool EmptyFields()
        {
            if (Globals.ThisDocument.txtOggetto.Text == "" || Globals.ThisDocument.txtDescrizione.Text == "")
            {
                MessageBox.Show("Alcuni campi obbligatori non sono stati compilati. Compilare i campi evidenziati!", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Error);

                Globals.ThisDocument.RemoveProtection();
                Globals.ThisDocument.Application.ScreenUpdating = false;

                ThisDocument.ToNormal("Oggetto", Word.WdColorIndex.wdBlack, "*");
                ThisDocument.ToNormal("Descrizione", Word.WdColorIndex.wdBlack, "*");

                if (Globals.ThisDocument.txtOggetto.Text == "")
                    ThisDocument.Highlight("Oggetto", Word.WdColorIndex.wdRed, "*");

                if (Globals.ThisDocument.txtDescrizione.Text == "")
                    ThisDocument.Highlight("Descrizione", Word.WdColorIndex.wdRed, "*");

                Globals.ThisDocument.Application.ScreenUpdating = true;
                Globals.ThisDocument.AddProtection();

                return true;
            }

            Globals.ThisDocument.RemoveProtection();
            Globals.ThisDocument.Application.ScreenUpdating = false;

            ThisDocument.ToNormal("Oggetto", Word.WdColorIndex.wdBlack, "*");
            ThisDocument.ToNormal("Descrizione", Word.WdColorIndex.wdBlack, "*");

            Globals.ThisDocument.Application.ScreenUpdating = true;
            Globals.ThisDocument.AddProtection();

            return false;
        }

        private string getAvailableID()
        {
            DataTable dt = _db.Select("spGetFirstAvailableID");
            return dt.Rows[0][0].ToString();
        }

        private void Print()
        {
            object missing = Missing.Value;

            if (Globals.ThisDocument.Application.Dialogs[Microsoft.Office.Interop.Word.WdWordDialog.wdDialogFilePrint].Show() == 1)
            {
                Globals.ThisDocument.PrintOut(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
        }

        private void ChangeBozzaVisibility(bool visible)
        {
            if (!visible)
            {
                Globals.ThisDocument.lbBozza.Text = "";
                Globals.ThisDocument.lbBozza.Image = null;
            }
            else
            {
                Globals.ThisDocument.lbBozza.Text = "Bozza";
                Globals.ThisDocument.lbBozza.Image = Resources.Editing_Edit_icon;
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

            DataTable dt = _db.Select("spGetAvailableYears");
            _cbAnniDispLabels = new List<string>();
            foreach (DataRow r in dt.Rows) 
            {
                RibbonDropDownItem i = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                i.Label = r["Anno"].ToString();
                _cbAnniDispLabels.Add(r["Anno"].ToString());
            }
            _cbAnniDispCount = _cbAnniDispLabels.Count;
            _appV = getCurrentV();
            _coreV = _db.GetCurrentV();
        }

        public void btnReset_Click(Office.IRibbonControl control)
        {
            if (MessageBox.Show("Sicuro di voler cancellare il contenuto dei campi?", "Cancellare campi?", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                Globals.ThisDocument.cmbStrumento.SelectedIndex = 0;
                ((DataView)Globals.ThisDocument.cmbStrumento.DataSource).RowFilter = "";
                Globals.ThisDocument.cmbStrumento.Enabled = true;
                Globals.ThisDocument.txtDescrizione.Text = "";
                Globals.ThisDocument.txtOggetto.Text = "";
                Globals.ThisDocument.txtNote.Text = "";
                Globals.ThisDocument.dtDataCreazione.Value = DateTime.Now;

                _btnRefreshEnabled = true;
                _btnSalvaBozzaEnabled = true;
                _chkIsDraft = false;
                this.ribbon.InvalidateControl("chkIsDraft");
                this.ribbon.Invalidate();
                getAvailableID();
            }
        }
        public void btnInvia_Click(Office.IRibbonControl control)
        {
            if (_chkIsDraft)
            {
                MessageBox.Show("La richiesta è contrassegnata come bozza. Togliere la spunta e riprovare.", "Impossibile salvare!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else
            {
                object oTrue = true;
                object oFalse = false;
                object missing = Missing.Value;

                QryParams parameters = new QryParams()
                {
                    {"@IdRichiesta", Globals.ThisDocument.lbIdRichiesta.Text}
                };

                DataView dv = _db.Select("spGetRichiesta", parameters).DefaultView;
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
                            Globals.ThisDocument.RemoveProtection();
                            Globals.ThisDocument.Application.ScreenUpdating = false;

                            ThisDocument.ToNormal("Oggetto", Word.WdColorIndex.wdBlack, "*");
                            ThisDocument.ToNormal("Descrizione", Word.WdColorIndex.wdBlack, "*");

                            Globals.ThisDocument.Application.ScreenUpdating = true;
                            Globals.ThisDocument.AddProtection();

                            _btnSalvaBozzaEnabled = false;

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
                                    {"@IdTipologiaStato", 1},
                                    {"@IdApplicazione", strumento["IdApplicazione"]},
                                    {"@Oggetto", Globals.ThisDocument.txtOggetto.Text.Trim()},
                                    {"@Descr", Globals.ThisDocument.txtDescrizione.Text.Trim()},
                                    {"@Note", Globals.ThisDocument.txtNote.Text.Trim()},
                                    {"@NomeFile", savePath}
                                };

                                _db.Insert("spSaveRichiestaModifica", parameters);
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
                _formAnnullaModifica = new FormAnnullaModifica(_cbAnniDispValue);
                _formAnnullaModifica.Show();
            }
            _formAnnullaModifica.WindowState = FormWindowState.Normal;
            _formAnnullaModifica.Focus();
        }
        public void btnSalvaBozza_Click(Office.IRibbonControl control)
        {
            if (!EmptyFields())
            {
                DateTime dataInvio = DateTime.Parse(Globals.ThisDocument.lbDataInvio.Text);
                DataRowView strumento = (DataRowView)Globals.ThisDocument.cmbStrumento.SelectedItem;

                _chkIsDraft = true;
                this.ribbon.InvalidateControl("chkIsDraft");

                QryParams parameters = new QryParams()
                {
                    {"@IdRichiesta", Globals.ThisDocument.lbIdRichiesta.Text},
                    {"@DataCreazione", Globals.ThisDocument.dtDataCreazione.Value.ToString("yyyyMMdd")},
                    {"@DataInvio", dataInvio.ToString("yyyyMMdd")},
                    {"@IdTipologiaStato", 7},
                    {"@IdApplicazione", strumento["IdApplicazione"]},
                    {"@Oggetto", Globals.ThisDocument.txtOggetto.Text.Trim()},
                    {"@Descr", Globals.ThisDocument.txtDescrizione.Text.Trim()},
                    {"@Note", Globals.ThisDocument.txtNote.Text.Trim()}
                };

                try
                {
                    _db.Insert("spSaveRichiestaModifica", parameters);
                }
                catch (Exception)
                {
                    MessageBox.Show("Salvataggio non riuscito... Riprovare più tardi.", "Errore!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        public void btnModifica_Click(Office.IRibbonControl control)
        {
            SelezionaModifica selMod = new SelezionaModifica(_cbAnniDispValue, _chkIsDraft, _btnRefreshEnabled);
            selMod.ShowDialog();
            _chkIsDraft = selMod._chkIsDraft;
            _btnRefreshEnabled = selMod._btnRefreshEnabled;
            this.ribbon.InvalidateControl("chkIsDraft");
            this.ribbon.Invalidate();
            selMod.Dispose();
        }
        public void chkIsDraft_Click(Office.IRibbonControl control, bool pressed)
        {
            ChangeBozzaVisibility(pressed);
            _chkIsDraft = pressed;
        }
        
        public bool chkIsDraft_getPressed(Office.IRibbonControl control) 
        {
            ChangeBozzaVisibility(_chkIsDraft);
            return _chkIsDraft;
        }
        
        public int cbAnniDisp_ItemCount(Office.IRibbonControl control)
        {
            return _cbAnniDispCount;
        }
        public string cbAnniDisp_ItemLabel(Office.IRibbonControl control, int i)
        {
            return _cbAnniDispLabels[i];
        }
        public int cbAnniDisp_getSelectedItemIndex(Office.IRibbonControl control)
        {
            return _cbAnniDispIndex;
        }
        public void cbAnniDisp_onAction(Office.IRibbonControl control, string itemID, int itemIndex)
        {
            _cbAnniDispValue = _cbAnniDispLabels[itemIndex];
            _cbAnniDispIndex = itemIndex;
        }
        public string lbVersioneApp_getLabel(Office.IRibbonControl control)
        {
            return "  App v" + _appV.ToString();
        }
        public string lbCoreV_getLabel(Office.IRibbonControl control)
        {
            return "  Core v" + _coreV.ToString();
        }
        public bool chkIsDraft_getEnabled(Office.IRibbonControl control)
        {
            return _chkIsDraftEnabled;
        }
        public bool btnSalvaBozza_enabled(Office.IRibbonControl control)
        {
            return _btnSalvaBozzaEnabled;
        }
        public bool btnRefresh_getEnabled(Office.IRibbonControl control)
        {
            return _btnRefreshEnabled;
        }
        
        public Bitmap btnReset_getImage(Office.IRibbonControl control)
        {
            return Resources.Reset_icon;
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
        public Bitmap btnSalvaBozza_getImage(Office.IRibbonControl control)
        {
            return Resources.save_icon;
        }
        public Bitmap btnModifica_getImage(Office.IRibbonControl control)
        {
            return Resources.edit_icon;
        }
        public Bitmap cbAnniDisponibili_getImage(Office.IRibbonControl control)
        {
            return Resources.calendar_icon;
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