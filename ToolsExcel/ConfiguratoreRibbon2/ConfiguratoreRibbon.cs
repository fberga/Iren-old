using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ConfiguratoreRibbon2
{
    public partial class ConfiguratoreRibbon : Form
    {
        const string NEW_GROUP_PREFIX = "New Group";
        const string NEW_BUTTON_PREFIX = "New Button";

        List<Panel> _groups = new List<Panel>();
        Panel _selectedPanel = new Panel();

        public ConfiguratoreRibbon()
        {
            InitializeComponent();
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

        private void CambioNomeGruppo(object sender, EventArgs e)
        {
            TextBox txt = sender as TextBox;

            if (txt.Name != txt.Text)
            {
                if (MessageBox.Show("Cambiare il nome del gruppo?", "Cambiare nome?", MessageBoxButtons.YesNo) == DialogResult.Yes) 
                {
                    txt.Name = txt.Text;
                }
                else
                    txt.Text = txt.Name;
            }

            UpdateGroupDimension(txt.Parent);

            txt.ReadOnly = true;
            txt.BackColor = txt.Parent.BackColor;
            txt.BorderStyle = BorderStyle.None;
        }

        private void DrawRightBorder(object sender, PaintEventArgs e)
        {
            Control group = sender as Control;
            Rectangle r = group.ClientRectangle;

            ControlPaint.DrawBorder3D(e.Graphics, r, Border3DStyle.Etched, Border3DSide.Right);
        }


        private SizeF MeasureTextSize(Control txt)
        {
            return txt.CreateGraphics().MeasureString(txt.Text, txt.Font, int.MaxValue);
        }

        private void AggiungiGruppo_Click(object sender, EventArgs e)
        {
            Panel newGroup = new Panel();

            panelRibbonLayout.Controls.Add(newGroup);
            newGroup.BringToFront();
            newGroup.Dock = DockStyle.Left;
            newGroup.BackColor = ControlPaint.LightLight(panelRibbonLayout.BackColor);
            newGroup.Padding = new Padding(0, 0, 2, 0);
            newGroup.Paint += DrawRightBorder;
            newGroup.Click += ChangeFocus;

            _selectedPanel = newGroup;

            TextBox txtNewGroup = new TextBox();
            newGroup.Controls.Add(txtNewGroup);
            txtNewGroup.Dock = DockStyle.Bottom;
            txtNewGroup.AutoSize = false;
            txtNewGroup.TextAlign = HorizontalAlignment.Center;


            int prog = FindLastOfItsKind(NEW_GROUP_PREFIX, typeof(TextBox)) + 1;

            txtNewGroup.Name = NEW_GROUP_PREFIX + " " + prog;
            txtNewGroup.Text = NEW_GROUP_PREFIX + " " + prog;

            newGroup.Width = (int)(MeasureTextSize(txtNewGroup).Width + 20);
            newGroup.Refresh();

            txtNewGroup.BorderStyle = BorderStyle.None;
            txtNewGroup.BackColor = ControlPaint.LightLight(panelRibbonLayout.BackColor);
            txtNewGroup.ReadOnly = true;
            
            txtNewGroup.Click += ChangeFocus;
            txtNewGroup.DoubleClick += EnterEditMode;
            txtNewGroup.KeyDown += ReturnPressed;
            txtNewGroup.LostFocus += CambioNomeGruppo;

            newGroup.ControlAdded += UpdateGroupDimension;
        }

        private void ChangeFocus(object sender, EventArgs e)
        {
            Control ctrl = sender as Control;
            ctrl.Select();
            _selectedPanel = GetAnchestorGroup(ctrl);
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
            Panel container = new Panel();
            container.Size = new Size(50, _selectedPanel.Height - 25);
            //container.BorderStyle = BorderStyle.FixedSingle;
            
            Button newBtn = new Button();
            
            container.Controls.Add(newBtn);
            newBtn.AutoSize = false;
            newBtn.MinimumSize = new Size(50, 100);

            newBtn.ImageAlign = ContentAlignment.TopCenter;
            newBtn.TextImageRelation = TextImageRelation.ImageAboveText;
            newBtn.TextAlign = ContentAlignment.MiddleCenter;
            newBtn.FlatStyle = FlatStyle.Flat;
            newBtn.FlatAppearance.BorderSize = 0;
            chooseImageDialog.InitialDirectory = @"D:\Repository\Iren\ToolsExcel\ToolsExcelBase\resources";
            chooseImageDialog.Filter = "PNG Files (*.png)|*.png";

            if (chooseImageDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                newBtn.Image = Image.FromFile(chooseImageDialog.FileName);
                //calcolo la posizione del tasto
                var left =
                    (from p in _selectedPanel.Controls.OfType<Panel>()
                     select p.Right).DefaultIfEmpty().Max();
                container.Left = left;
                _selectedPanel.Controls.Add(container);

                container.SizeChanged += UpdateParentGroupDimension;
                newBtn.SizeChanged += UpdateContainerWidth;
                newBtn.Click += OpenButtonCfgForm;

                int prog = FindLastOfItsKind(NEW_BUTTON_PREFIX, typeof(Button)) + 1;
                newBtn.Name = NEW_BUTTON_PREFIX + " " + prog;

                newBtn.Text = NEW_BUTTON_PREFIX + " " + prog;
                newBtn.Tag = new Dictionary<string, object>();

                newBtn.Width = Math.Min((int)(MeasureTextSize(newBtn).Width + 4), 250);

            }
        }

        private void UpdateContainerWidth(object sender, EventArgs e)
        {
            Control ctrl = sender as Control;
            
            Control parent = ctrl.Parent;
            var maxBtnWidth =
                (from btn in parent.Controls.OfType<Button>()
                 select btn.Width).DefaultIfEmpty().Max();
            var maxCmbWidth =
                (from cmb in parent.Controls.OfType<ComboBox>()
                 select cmb.Width).DefaultIfEmpty().Max();
            var maxLabelWidth =
                (from lbl in parent.Controls.OfType<Label>()
                 select (int)MeasureTextSize(lbl).Width).DefaultIfEmpty().Max();

            parent.Width = Math.Max(Math.Max(maxBtnWidth, maxCmbWidth), maxLabelWidth);
        }

        private void OpenButtonCfgForm(object sender, EventArgs e)
        {
            ConfiguratoreTasto cfg = new ConfiguratoreTasto(sender as Button);
            cfg.Show();
        }

        private void ChangePosition(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        private void UpdateGroupDimension(Control parent)
        {
            var txtWidth =
                (from txt in parent.Controls.OfType<TextBox>()
                 select (int)(MeasureTextSize(txt).Width + 20)).FirstOrDefault();

            var totWidth =
                (from p in parent.Controls.OfType<Panel>()
                 select p.Width).DefaultIfEmpty().Sum() + 20;

            parent.Width = Math.Max(txtWidth, totWidth);
            parent.Invalidate();
        }

        private void UpdateGroupDimension(object sender, ControlEventArgs e)
        {
            UpdateGroupDimension(sender as Control);
        }

        private void UpdateParentGroupDimension(object sender, EventArgs e)
        {
            Control ctrl = sender as Control;
            UpdateGroupDimension(ctrl.Parent);
        }

        private void BtnImageChosen(object sender, CancelEventArgs e)
        {
        }
    }
}
