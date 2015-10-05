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

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class ConfiguratoreRibbon : Form
    {
        public const string IMG_DIRECTORY = @"D:\Repository\Iren\ToolsExcel\ToolsExcelBase\resources";

        private const string APPLICAZIONI = "spApplicazioneProprieta";

        private string[] _ambienti;

        public ConfiguratoreRibbon()
        {            
            InitializeComponent();
            string[] files = Directory.GetFiles(IMG_DIRECTORY);
            var imgs =
                from f in files
                where Regex.IsMatch(f, @".+\.(?:png|jpg|bmp)")
                select f;

            foreach (string img in imgs)
            {
                imageListNormal.Images.Add(img, Image.FromFile(img));
                imageListSmall.Images.Add(img, Image.FromFile(img));
            }

            //inizializzazione connessione
            _ambienti = Workbook.AppSettings("Ambienti").Split('|');
            DataBase.InitNewDB(_ambienti[0]);

            //carico la lista di applicazioni configurabili
            DataTable applicazioni = DataBase.Select(APPLICAZIONI, "@IdApplicazione=0");
            if (applicazioni != null)
            {
                cmbApplicazioni.DisplayMember = "DesApplicazione";
                cmbApplicazioni.ValueMember = "IdApplicazione";
                cmbApplicazioni.DataSource = applicazioni;
            }

        }       

        private Panel GetAnchestorGroup(Control ctrl)
        {
            if (ctrl.GetType() == typeof(Panel) && ctrl.Parent == panelRibbonLayout)
                return ctrl as Panel;

            if (ctrl == panelRibbonLayout)
                return null;

            if (ctrl == this)
                return null;

            return GetAnchestorGroup(ctrl.Parent);
        }

        private void UpdateParentGroupDimension(object sender, EventArgs e)
        {
            Control ctrl = sender as Control;
            Utility.UpdateGroupDimension(ctrl.Parent);
        }

        private void ctrlDownButton_Click(object sender, EventArgs e)
        {
            IRibbonComponent ctrl = ActiveControl as IRibbonComponent;

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
        private void ctrlUpButton_Click(object sender, EventArgs e)
        {
            IRibbonComponent ctrl = ActiveControl as IRibbonComponent;

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
        private void ctrlLeftButton_Click(object sender, EventArgs e)
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
        private void ctrlRightButton_Click(object sender, EventArgs e)
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

            ControlContainer container = new ControlContainer();
            container.Size = new Size(50, ActiveControl.Height - 30);

            var left =
                (from p in ActiveControl.Controls.OfType<ControlContainer>()
                 select p.Right).DefaultIfEmpty().Max();

            container.Left = left == 0 ? ActiveControl.Padding.Left : left + 10;
            container.Top = ActiveControl.Padding.Top;
            return container;
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
    }
}
