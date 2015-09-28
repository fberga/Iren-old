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

namespace ConfiguratoreRibbon2
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

        private IEnumerable<Control> GetAll(Control control, Type type)
        {
            var controls = control.Controls.Cast<Control>();

            return controls.SelectMany(ctrl => GetAll(ctrl, type))
                                      .Concat(controls)
                                      .Where(c => c.GetType() == type);
        }

        private int FindLastOfItsKind(string prefix, Type type)
        {
            var progs = GetAll(panelRibbonLayout, type)
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

        private void ReturnPressed(object sender, KeyEventArgs e)
        {
            TextBox txt = sender as TextBox;
            if (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab)
            {
                e.SuppressKeyPress = true;
                txt.ReadOnly = true;
                txt.Parent.Focus();
            }
        }

        private void EnterEditMode(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;
            txt.ReadOnly = false;
            txt.SelectAll();
        }

        private void AggiungiGruppo_Click(object sender, EventArgs e)
        {
            GroupPanel newGroup = new GroupPanel(panelRibbonLayout);
            panelRibbonLayout.Controls.Add(newGroup);
            newGroup.BringToFront();
            newGroup.Select();
            //_selectedGroup = newGroup;
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
            if(ActiveControl.GetType() != typeof(GroupPanel)) 
            {
                MessageBox.Show("Nessun gruppo selezionato...", "ATTENZIONE!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            ControlContainer container = new ControlContainer();
            container.Size = new Size(50, ActiveControl.Height - 30);
            RibbonButton newBtn = new RibbonButton(imageListNormal, imageListSmall);

            var left =
                (from p in ActiveControl.Controls.OfType<ControlContainer>()
                 select p.Right).DefaultIfEmpty().Max();

            container.Left = left == 0 ? ActiveControl.Padding.Left : left;
            container.Top = ActiveControl.Padding.Top;

            container.Controls.Add(newBtn);
        }

        private void UpdateParentGroupDimension(object sender, EventArgs e)
        {
            Control ctrl = sender as Control;
            Utility.UpdateGroupDimension(ctrl.Parent);
        }
    }
}
