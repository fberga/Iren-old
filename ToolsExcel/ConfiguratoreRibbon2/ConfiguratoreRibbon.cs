using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Iren.ToolsExcel.Utility;
using Iren.ToolsExcel.Base;
using System.Globalization;
using System.Collections;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class ConfiguratoreRibbon : Form
    {
        private string[] _ambienti;

        //public static int IdApplicazione { get; private set; }
        public static List<int> IdUtenti { get { return new List<int>() { 62 }; } }// private set; }

        public static List<int> ControlliUtilizzati { get; private set; }
        public static List<int> GruppiUtilizzati { get; private set; }
        public static List<int> GruppoControlloUtilizzati { get; private set; }
        public static List<int> GruppoControlloCancellati { get; set; }

        public ConfiguratoreRibbon()
        {
            Utility.InitializeUtility();
            Utility.StdFont = this.Font;
            InitializeComponent();

            //trovo tutte le risorse disponibili in Iren.ToolsExcel.Base
            var resourceSet = Iren.ToolsExcel.Base.Properties.Resources.ResourceManager.GetResourceSet(CultureInfo.InstalledUICulture, true, true);
            
            //Considero solo quelle che sono di tipo Image
            var imgs =
                from r in resourceSet.Cast<DictionaryEntry>()
                where r.Value is Image
                select r;

            foreach (var img in imgs)
            {
                Utility.ImageListNormal.Images.Add(img.Key as string, img.Value as Image);
                Utility.ImageListSmall.Images.Add(img.Key as string, img.Value as Image);
            }

            //inizializzazione connessione
            _ambienti = Workbook.AppSettings("Ambienti").Split('|');
            DataBase.InitNewDB(_ambienti[0]);

            //carico la lista di applicazioni configurabili
            DataTable applicazioni = DataBase.Select(SP.APPLICAZIONI, "@IdApplicazione=0");
            if (applicazioni != null)
            {
                cmbApplicazioni.DisplayMember = "DesApplicazione";
                cmbApplicazioni.ValueMember = "IdApplicazione";
                cmbApplicazioni.DataSource = applicazioni;
            }

            //carico la lista degli utenti
            DataTable utenti = DataBase.Select(SP.UTENTI, "@IdUtenteGruppo=5");
            if (utenti != null)
            {
                cmbUtenti.DisplayMember = "Nome";
                cmbUtenti.ValueMember = "IdUtente";
                cmbUtenti.DataSource = utenti;
            }
        }

        private void CaricaAnteprimaRibbon()
        {
            ControlliUtilizzati = new List<int>();
            GruppiUtilizzati = new List<int>();
            GruppoControlloUtilizzati = new List<int>();
            GruppoControlloCancellati = new List<int>();

            DataTable ribbon = DataBase.Select(SP.GRUPPO_CONTROLLO);
            DataTable funzioni = DataBase.Select(SP.CONTROLLO_FUNZIONE);

            if (ribbon != null)
            {
                int idGroup = -1;
                RibbonGroup grp = null;
                foreach (DataRow r in ribbon.Rows)
                {
                    //prendo nota di cosa è utilizzato.
                    GruppoControlloUtilizzati.Add((int)r["IdGruppoControllo"]);
                    ControlliUtilizzati.Add((int)r["IdControllo"]);
                    GruppiUtilizzati.Add((int)r["IdGruppo"]);

                    if (!r["IdGruppo"].Equals(idGroup))
                    {
                        idGroup = (int)r["IdGruppo"];
                        grp = new RibbonGroup(panelRibbonLayout, (int)r["IdGruppo"]);
                        panelRibbonLayout.Controls.Add(grp);
                        grp.Text = r["LabelGruppo"].ToString();
                    }
                    
                    Control ctrl = Utility.AddControlToGroup(grp, r, funzioni);
                    ctrl.ContextMenuStrip = new DisabilitaMenuStrip();
                }
            }
        }

        //SPOSTAMENTI
        private void MoveDown_Click(object sender, EventArgs e)
        {
            IRibbonControl ctrl = ActiveControl as IRibbonControl;

            if (ctrl != null && ctrl.Slot < 3)
            {
                var nextCtrl = Utility.GetAll(ActiveControl.Parent)
                        .Where(c => ActiveControl.Bottom == c.Top).FirstOrDefault();

                if (nextCtrl != null)
                {
                    nextCtrl.Top = ActiveControl.Top;
                    ActiveControl.Top = nextCtrl.Bottom;
                }
            }
        }
        private void MoveUp_Click(object sender, EventArgs e)
        {
            IRibbonControl ctrl = ActiveControl as IRibbonControl;

            if (ctrl != null && ctrl.Slot < 3)
            {
                var nextCtrl = Utility.GetAll(ActiveControl.Parent)
                        .Where(c => ActiveControl.Top == c.Bottom).FirstOrDefault();

                if (nextCtrl != null)
                {
                    ActiveControl.Top = nextCtrl.Top;
                    nextCtrl.Top = ActiveControl.Bottom;
                }
            }
        }
        private void MoveLeft_Click(object sender, EventArgs e)
        {
            if(ActiveControl != null) 
            {
                var oth = ActiveControl.Parent.Controls.Cast<Control>()
                    .Where(c => c.Right == ActiveControl.Left)
                    .FirstOrDefault();

                if(oth != null)
                {
                    ActiveControl.Left = oth.Left;
                    oth.Left = ActiveControl.Right;
                }
            }
        }
        private void MoveRight_Click(object sender, EventArgs e)
        {
            if (ActiveControl != null)
            {
                var oth = ActiveControl.Parent.Controls.Cast<Control>()
                    .Where(c => c.Left == ActiveControl.Right)
                    .FirstOrDefault();

                if (oth != null)
                {
                    oth.Left = ActiveControl.Left;
                    ActiveControl.Left = oth.Right;
                }
            }
        }

        private bool IsRibbonGroupSelected()
        {
            if (ActiveControl.GetType() != typeof(RibbonGroup))
            {
                MessageBox.Show("Nessun gruppo selezionato...", "ATTENZIONE!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private Control CreateEmptyContainer()
        {
            if (IsRibbonGroupSelected())
                return Utility.CreateEmptyContainer(ActiveControl);

            return null;   
        }

        //TASTI
        private void AggiungiNuovoTasto_Click(object sender, EventArgs e)
        {
            Control container = CreateEmptyContainer();
            if (container != null)
            {
                RibbonButton newBtn = new RibbonButton(panelRibbonLayout);
                if (newBtn.ImageKey != "")
                {
                    container.Controls.Add(newBtn);
                    newBtn.Top = container.Padding.Top;
                    newBtn.Left = container.Padding.Left;
                    ActiveControl.Controls.Add(container);
                }
            }
        }
        private void ScegliTastoEsistente_Click(object sender, EventArgs e)
        {
            if (IsRibbonGroupSelected())
            {
                using (ControlliEsistenti ctrlForm = new ControlliEsistenti(ActiveControl, 1, 2))
                {
                    ctrlForm.ShowDialog();
                }
            }
        }
        
        //COMBO
        private void AggiungiNuovoCombo_Click(object sender, EventArgs e)
        {
            Control container = CreateEmptyContainer();
            if (container != null)
            {
                RibbonComboBox newDrpDwn = new RibbonComboBox(panelRibbonLayout);
                if(!newDrpDwn.IsDisposed)
                {
                    container.Controls.Add(newDrpDwn);
                    newDrpDwn.Top = container.Padding.Top;
                    newDrpDwn.Left = container.Padding.Left;

                    ActiveControl.Controls.Add(container);
                }
            }
        }
        private void ScegliComboEsistente_Click(object sender, EventArgs e)
        {
            if (IsRibbonGroupSelected())
            {
                using (ControlliEsistenti ctrlForm = new ControlliEsistenti(ActiveControl, 3))
                {
                    ctrlForm.ShowDialog();
                }
            }
        }
        

        //GRUPPI
        private void AggiungiGruppo_Click(object sender, EventArgs e)
        {
            RibbonGroup newGroup = new RibbonGroup(panelRibbonLayout);
            if(!newGroup.IsDisposed)
                Utility.AddGroupToRibbon(panelRibbonLayout, newGroup);
        }
        private void ScegliGruppoEsistente_Click(object sender, EventArgs e)
        {
            using (GruppiEsistenti grpForm = new GruppiEsistenti(panelRibbonLayout))
            {
                grpForm.ShowDialog();
            }
        }

        //VUOTI
        private void AggiungiContenitoreVuoto_Click(object sender, EventArgs e)
        {
            Control container = CreateEmptyContainer();
            if (container != null)
                ActiveControl.Controls.Add(container);
        }

        private void ApplicaConfigurazione_Click(object sender, EventArgs e)
        {
            //rimuovo i componenti cancellati
            if(GruppoControlloCancellati.Count > 0) 
            {
                DataBase.Delete(SP.DELETE_GRUPPO_CONTROLLO, "@Ids=" + string.Join(",", GruppoControlloCancellati));
            }
            //trovo tutti i gruppi
            var groups =
                panelRibbonLayout.Controls.OfType<RibbonGroup>().OrderBy(g => g.Left);
            
            int ordine = 1;
            foreach (RibbonGroup group in groups)
            {
                //trovo tutti i contenitori
                var containers =
                    group.Controls.OfType<ControlContainer>().OrderBy(c => c.Left);

                Dictionary<string, object> outP = new Dictionary<string,object>();
                int groupId = -1;
                if (DataBase.Insert(SP.INSERT_GRUPPO, new Core.QryParams()
                    {
                        {"@Id", group.ID},
                        {"@Nome", group.Name},
                        {"@Label", group.Text}
                    }, out outP))
                    groupId = (int)outP["@Id"];

                foreach (var container in containers)
                {
                    //trovo tutti i controlli contenuti nei contenitori
                    var ctrls =
                        container.Controls.Cast<IRibbonControl>();

                    foreach (IRibbonControl ctrl in ctrls)
                    {
                        outP = new Dictionary<string, object>();
                        int ctrlId = -1;
                        if (DataBase.Insert(SP.INSERT_CONTROLLO, new Core.QryParams()
                            {
                                {"@Id", ctrl.IdControllo},
                                {"@IdTipologiaControllo", ctrl.IdTipologia},
                                {"@Nome", ctrl.Name},
                                {"@Descrizione", ctrl.Description},
                                {"@Immagine", ctrl.ImageKey},
                                {"@Label", ctrl.Text},
                                {"@ScreenTip", ctrl.ScreenTip},
                                {"@ControlSize", ctrl.Dimension}
                            }, out outP))
                            ctrlId = (int)outP["@Id"];

                        if(DataBase.Insert(SP.INSERT_GRUPPO_CONTROLLO, new Core.QryParams() { 
                            {"@Id", 0},
                            {"@IdApplicazione", 1},
	                        {"@IdUtente", 62},
	                        {"@IdGruppo", groupId},
	                        {"@IdControllo", ctrlId},
                            {"@Abilitato", ctrl.Enabled ? "1" : "0"},
	                        {"@Ordine", ordine++}
                        }, out outP))
                        {
                            int ordineFunzioni = 1;
                            foreach(int idFunzione in ctrl.Functions)
                            {
                                DataBase.Insert(SP.INSERT_CONTROLLO_FUNZIONE, new Core.QryParams()
                                {
                                    {"@IdGruppoControllo", outP["@Id"]},
                                    {"@IdFunzione", idFunzione},
                                    {"@Ordine", ordineFunzioni++},
                                });
                            }
                        }
                    }
                }
            }
            //refresh dell'anteprima
            RicaricaRibbon_Click(null, null);
        }
        private void RicaricaRibbon_Click(object sender, EventArgs e)
        {
            panelRibbonLayout.Controls.Clear();
            CaricaAnteprimaRibbon();
        }

        private void CambioApplicazione(object sender, EventArgs e)
        {
            if (cmbApplicazioni.SelectedValue != null)
                DataBase.DB.SetParameters(idApplicazione: (int)cmbApplicazioni.SelectedValue);

            RicaricaRibbon_Click(null, null);
        }

        private void CambioUtente(object sender, EventArgs e)
        {
            if(cmbUtenti.SelectedValue != null)
                DataBase.DB.SetParameters(idUtente: (int)cmbUtenti.SelectedValue);

            RicaricaRibbon_Click(null, null);
        }
    }
}
