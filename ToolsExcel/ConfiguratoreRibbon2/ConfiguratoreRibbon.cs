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

        public ConfiguratoreRibbon()
        {
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
                imageListNormal.Images.Add(img.Key as string, img.Value as Image);
                imageListSmall.Images.Add(img.Key as string, img.Value as Image);
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
            //carico l'eventiale ribbon già definita
            CaricaAnteprimaRibbon(1, 62);
        }

        private void CaricaAnteprimaRibbon(int appID, int usrID)
        {
            DataTable ribbon = DataBase.Select(SP.APPLICAZIONE_UTENTE_RIBBON, "@IdApplicazione=" + appID + ";@IdUtente=" + usrID);

            string group = "";
            RibbonGroup grp = null;
            foreach (DataRow r in ribbon.Rows)
            {
                if (!r["NomeGruppo"].Equals(group))
                {
                    group = r["NomeGruppo"].ToString();
                    
                    grp = new RibbonGroup(panelRibbonLayout, (int)r["IdGruppo"]);                    
                    panelRibbonLayout.Controls.Add(grp);
                    grp.Label = r["LabelGruppo"].ToString();
                }
                
                Control container = Utility.CreateEmptyContainer(grp);
                switch((int)r["IdTipologiaControllo"]) 
                {
                    case 1:
                    case 2:                       
                        grp.Controls.Add(container);

                        RibbonButton btn = new RibbonButton(imageListNormal, imageListSmall, r["Immagine"].ToString(), (int)r["IdControllo"]);
                        container.Controls.Add(btn);

                        btn.Top = container.Padding.Top;
                        btn.Left = container.Padding.Left;

                        btn.Descrizione = r["Descrizione"].ToString();
                        btn.Label = r["Label"].ToString();
                        btn.ScreenTip = r["ScreenTip"].ToString();
                        btn.Dimensione = (int)r["ControlSize"];
                        btn.ToggleButton = r["IdTipologiaControllo"].Equals(2);

                        break;
                    case 3:                        
                        grp.Controls.Add(container);

                        RibbonDropDown drpD = new RibbonDropDown((int)r["IdControllo"]);
                        container.Controls.Add(drpD);

                        drpD.Top = container.Padding.Top;
                        drpD.Left = container.Padding.Left;

                        drpD.Descrizione = r["Descrizione"].ToString();
                        drpD.Label = r["Label"].ToString();
                        drpD.SetWidth();
                        drpD.ScreenTip = r["ScreenTip"].ToString();

                        break;
                }
            }
        }

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

        
        private Control CreateEmptyContainer()
        {
            if (ActiveControl.GetType() != typeof(RibbonGroup))
            {
                MessageBox.Show("Nessun gruppo selezionato...", "ATTENZIONE!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            return Utility.CreateEmptyContainer(ActiveControl);   
        }

        private void AggiungiTasto_Click(object sender, EventArgs e)
        {
            Control container = CreateEmptyContainer();
            if (container != null)
            {
                RibbonButton newBtn = new RibbonButton(panelRibbonLayout, imageListNormal, imageListSmall);
                if (newBtn.ImageKey != "")
                {
                    container.Controls.Add(newBtn);
                    newBtn.Top = container.Padding.Top;
                    newBtn.Left = container.Padding.Left;
                    ActiveControl.Controls.Add(container);
                }
            }
        }
        private void AggiungiDropDown_Click(object sender, EventArgs e)
        {
            Control container = CreateEmptyContainer();
            if (container != null)
            {
                RibbonDropDown newDrpDwn = new RibbonDropDown(panelRibbonLayout);
                container.Controls.Add(newDrpDwn);
                newDrpDwn.Top = container.Padding.Top;
                newDrpDwn.Left = container.Padding.Left;

                ActiveControl.Controls.Add(container);
            }
        }
        private void AggiungiContenitoreVuoto_Click(object sender, EventArgs e)
        {
            Control container = CreateEmptyContainer();
            if(container != null)
                ActiveControl.Controls.Add(container);
        }
        private void AggiungiGruppo_Click(object sender, EventArgs e)
        {
            RibbonGroup newGroup = new RibbonGroup(panelRibbonLayout);
            int left = panelRibbonLayout.Controls.OfType<RibbonGroup>()
                .Select(c => c.Right)
                .DefaultIfEmpty()
                .Max();

            newGroup.Left = left == 0 ? panelRibbonLayout.Padding.Left : left;
            panelRibbonLayout.Controls.Add(newGroup);


            //newGroup.BringToFront();
            newGroup.Select();
        }

        private void btnSalva_Click(object sender, EventArgs e)
        {
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
                        {"@Label", group.Label}
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
                                {"@Id", ctrl.ID},
                                {"@IdTipologiaControllo", ctrl.IdTipologia},
                                {"@Descrizione", ctrl.Descrizione ?? ""},
                                {"@Immagine", ctrl.ImageName ?? ""},
                                {"@Label", ctrl.Label ?? ""},
                                {"@ScreenTip", ctrl.ScreenTip ?? ""},
                                {"@ControlSize", ctrl.Dimensione}
                            }, out outP))
                            ctrlId = (int)outP["@Id"];

                        DataBase.Insert(SP.INSERT_GRUPPO_CONTROLLO, new Core.QryParams() { 
                            {"@IdApplicazione", 1},
	                        {"@IdUtente", 62},
	                        {"@IdGruppo", groupId},
	                        {"@IdControllo", ctrlId},
	                        {"@Ordine", ordine++}
                        });
                    }
                }
            }
        }

        private void scegliToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ControlliEsistenti ctrlForm = new ControlliEsistenti(imageListSmall, imageListNormal, 1, 2);

            ctrlForm.Show();
        }        
    }
}
