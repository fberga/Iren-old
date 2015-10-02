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

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public partial class ConfiguratoreRibbon : Form
    {
        public const string IMG_DIRECTORY = @"D:\Repository\Iren\ToolsExcel\ToolsExcelBase\resources";

        public ConfiguratoreRibbon()
        {
            InitializeComponent();
            //trovo tutti i file contenuti
            string[] files = Directory.GetFiles(IMG_DIRECTORY);
            //filtro solo le immagini
            var imgs =
                from f in files
                where Regex.IsMatch(f, @".+\.(?:png|jpg|bmp)")
                select f;

            foreach (string img in imgs)
            {
                imageListNormal.Images.Add(img, Image.FromFile(img));
                imageListSmall.Images.Add(img, Image.FromFile(img));
            }

        }

        //private IEnumerable<Control> GetAll(Control control, Type type)
        //{
        //    var controls = control.Controls.Cast<Control>();

        //    return controls.SelectMany(ctrl => GetAll(ctrl, type))
        //                              .Concat(controls)
        //                              .Where(c => c.GetType() == type);
        //}

        private int FindLastOfItsKind(string prefix, Type type)
        {
            var progs = Utility.GetAll(panelRibbonLayout, type)
                .Where(c => c.Name.StartsWith(prefix))
                .Select(c => 
                    {
                        string num = Regex.Match(c.Name, @"\d+").Value;
                        int progNum = 0;
                        int.TryParse(num, out progNum);
                        return progNum;
                    }).ToList();

            if (progs.Count > 0)
                return progs.Max();

            return 0;
        }

        private void AggiungiGruppo_Click(object sender, EventArgs e)
        {
            RibbonGroup newGroup = new RibbonGroup(panelRibbonLayout);
            panelRibbonLayout.Controls.Add(newGroup);
            newGroup.BringToFront();
            newGroup.Select();
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

        private void AggiungiTasto_Click(object sender, EventArgs e)
        {
            if(ActiveControl.GetType() != typeof(RibbonGroup)) 
            {
                MessageBox.Show("Nessun gruppo selezionato...", "ATTENZIONE!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ControlContainer container = new ControlContainer();
            container.Size = new Size(50, ActiveControl.Height - 30);
            RibbonButton newBtn = new RibbonButton(panelRibbonLayout, imageListNormal, imageListSmall);

            var left =
                (from p in ActiveControl.Controls.OfType<ControlContainer>()
                 select p.Right).DefaultIfEmpty().Max();

            container.Left = left == 0 ? ActiveControl.Padding.Left : left;
            container.Top = ActiveControl.Padding.Top;
            if (newBtn.ImageKey != "")
            {
                ActiveControl.Controls.Add(container);
                container.Controls.Add(newBtn);
                newBtn.Top = container.Padding.Top;
                newBtn.Left = container.Padding.Left;
            }
        }

        private void UpdateParentGroupDimension(object sender, EventArgs e)
        {
            Control ctrl = sender as Control;
            Utility.UpdateGroupDimension(ctrl.Parent);
        }

        private void ctrlDownButton_Click(object sender, EventArgs e)
        {
            if (ActiveControl.GetType() == typeof(RibbonButton))
            {
                RibbonButton btn = ActiveControl as RibbonButton;

                if (btn.Slot== 1)
                {
                    var nextBtn = Utility.GetAll(btn.Parent)
                        .Where(c => btn.Bottom == c.Top).FirstOrDefault();

                    if (nextBtn != null)
                    {
                        nextBtn.Top = btn.Top;
                        btn.Top = nextBtn.Bottom;
                    }
                }
            }
        }

        private void ctrlUpButton_Click(object sender, EventArgs e)
        {
            if (ActiveControl.GetType() == typeof(RibbonButton))
            {
                RibbonButton btn = ActiveControl as RibbonButton;

                if (btn.Slot == 1)
                {
                    var nextBtn = Utility.GetAll(btn.Parent)
                        .Where(c => btn.Top == c.Bottom).FirstOrDefault();

                    if (nextBtn != null)
                    {
                        btn.Top = nextBtn.Top;
                        nextBtn.Top = btn.Bottom;
                    }
                }
            }
        }

        private void ctrlLeftButton_Click(object sender, EventArgs e)
        {
            //if(ActiveControl != null) 
            //{
            //    if(ActiveControl.Parent.GetType() == typeof(ControlContainer))
            //    {
            //        ControlContainer actual = ActiveControl.Parent as ControlContainer;
            //        var next =
            //            Utility.GetAll(panelRibbonLayout, typeof(ControlContainer))
            //            .Cast<ControlContainer>()
            //            .Where(c => c.Right == actual.Left && c.FreeSlot > )
            //    }
            //}
        }

        private void AddDropDown_Click(object sender, EventArgs e)
        {
            if (ActiveControl.GetType() != typeof(RibbonGroup))
            {
                MessageBox.Show("Nessun gruppo selezionato...", "ATTENZIONE!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ControlContainer container = new ControlContainer();
            container.Size = new Size(50, ActiveControl.Height - 30);
            RibbonDropDown newDrpDwn = new RibbonDropDown(panelRibbonLayout);

            var left =
                (from p in ActiveControl.Controls.OfType<ControlContainer>()
                 select p.Right).DefaultIfEmpty().Max();

            container.Left = left == 0 ? ActiveControl.Padding.Left : left;
            container.Top = ActiveControl.Padding.Top;
            ActiveControl.Controls.Add(container);
            container.Controls.Add(newDrpDwn);
            newDrpDwn.Top = container.Padding.Top;
            newDrpDwn.Left = container.Padding.Left;
        }
    }
}
