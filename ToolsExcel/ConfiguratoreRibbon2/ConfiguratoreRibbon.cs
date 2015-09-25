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
        const string NEW_GROUP_PREFIX = "New Group";
        const string NEW_BUTTON_PREFIX = "New Button";

        public const string DESC_FIELD_NAME = "Description";
        public const string SCREEN_TIP_FIELD_NAME = "ScreenTip";
        public const string TOGGLE_BUTTON_FIELD_NAME = "ToggleButton";
        public const string DIMENSION_FIELD_NAME = "Dimension";

        public const string IMG_DIRECTORY = @"D:\Repository\Iren\ToolsExcel\ToolsExcelBase\resources";

        Size largeBtnMinSize = new Size(50, 100);
        Size smallBtnMaxSize = new Size(250, 33);

        List<Panel> _groups = new List<Panel>();
        Panel _selectedPanel = new Panel();

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


        private SizeF MeasureTextSize(Control ctrl)
        {
            //calcolo la dimensione
            //lavoro su 2 righe...quindi calcolo tutte le dimensioni delle parole e poi le combino per tentativi mettendo:
            // 1 sopra, tot - 1 sotto; 2 sopra, tot - 2 sotto; ... 

            string s = ctrl.Text;
            
            //se è un tasto a dimensione piccola, calcolo normalmente
            object dim = null;
            if (ctrl.GetType() == typeof(Button) && ctrl.Tag != null)
            {
                Dictionary<string, object> metaData = ctrl.Tag as Dictionary<string, object>;
                metaData.TryGetValue(DIMENSION_FIELD_NAME, out dim);
            }

            if(!s.Contains(' ') || ctrl.GetType() == typeof(TextBox) || (int)(dim ?? 1) == 0)
                return ctrl.CreateGraphics().MeasureString(s, ctrl.Font, int.MaxValue);

            string[] parole = s.Split(' ');
            float[] misure = new float[parole.Length];

            //calcolo le singole dimensioni
            for (int i = 0; i < parole.Length; i++)
                misure[i] = ctrl.CreateGraphics().MeasureString(parole[i], ctrl.Font, int.MaxValue).Width;

            //provo a combinare tutte le parole e vedo quale combinazione mi da dimensione minima (forse anche rapporto più bilanciato...)
            float riga1 = Enumerable.Sum(misure);
            float riga2 = 0;

            //float rapporto = 0;
            float opt = riga1;

            //ciclo ma lascio almeno una parole sopra
            for (int i = parole.Length - 1; i > 0; i--)
            {
                riga2 += misure[i];
                riga1 -= misure[i];

                float tmpOpt = Math.Max(riga1, riga2);

                if (opt > tmpOpt) 
                {
                    opt = tmpOpt;
                }
            }

            return ctrl.CreateGraphics().MeasureString(s, ctrl.Font, (int)Math.Ceiling(opt));

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

        private void SetUpLargeButton(Button btn)
        {
            btn.ImageList = imageListNormal;
            btn.MaximumSize = new Size(int.MaxValue, int.MaxValue);
            btn.MinimumSize = largeBtnMinSize;
            btn.ImageAlign = ContentAlignment.TopCenter;
            btn.TextImageRelation = TextImageRelation.ImageAboveText;
            btn.TextAlign = ContentAlignment.MiddleCenter;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
        }
        private void SetUpSmallButton(Button btn)
        {
            btn.ImageList = imageListSmall;
            btn.MinimumSize = new Size(0, 0);
            btn.MaximumSize = smallBtnMaxSize;
            btn.ImageAlign = ContentAlignment.MiddleLeft;
            btn.TextImageRelation = TextImageRelation.ImageBeforeText;
            btn.TextAlign = ContentAlignment.MiddleLeft;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.AutoEllipsis = true;
        }

        private int CalculateLargeButtonWidth(Button btn)
        {
            return Math.Min((int)(MeasureTextSize(btn).Width + 15), 250);
        }
        private int CalculateSmallButtonWidth(Button btn)
        {
            return Math.Min((int)(MeasureTextSize(btn).Width + 30), 250);
        }

        private void AggiungiTasto_Click(object sender, EventArgs e)
        {
            Panel container = new Panel();
            container.Size = new Size(50, _selectedPanel.Height - 25);
            //TODO Rimuovere bordo
            container.BorderStyle = BorderStyle.FixedSingle;
            
            Button newBtn = new Button();
            SetUpLargeButton(newBtn);
            container.Controls.Add(newBtn);

            //SelettoreImmagini chooseImageDialog = new SelettoreImmagini(imageListNormal);
            using (SelettoreImmagini chooseImageDialog = new SelettoreImmagini(imageListNormal))
            {
                if (chooseImageDialog.ShowDialog() == DialogResult.OK)
                {
                    newBtn.ImageKey = chooseImageDialog.FileName;
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

                    newBtn.Tag = new Dictionary<string, object>() { {DIMENSION_FIELD_NAME, 1} };

                    newBtn.Width = CalculateLargeButtonWidth(newBtn);
                }
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

            parent.Width = Enumerable.Max(new int[]{maxBtnWidth, maxCmbWidth, maxLabelWidth});
        }

        private void OpenButtonCfgForm(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            using (ConfiguratoreTasto cfg = new ConfiguratoreTasto(btn, imageListNormal))
            {
                cfg.ShowDialog();

                //controllo se è cambiato il campo size (da grande a piccolo o viceversa) e adeguo il form di conseguenza
                Dictionary<string, object> metaData = btn.Tag as Dictionary<string, object>;
                object size;
                metaData.TryGetValue(DIMENSION_FIELD_NAME, out size);
                int value = (int)(size ?? 1);

                if (value == 1)
                {
                    SetUpLargeButton(btn);
                    btn.Width = CalculateLargeButtonWidth(btn);
                    btn.Height = btn.MinimumSize.Height;
                }
                else if (value == 0)
                {
                    SetUpSmallButton(btn);
                    btn.Width = CalculateSmallButtonWidth(btn);
                    //btn.Height = btn.MaximumSize.Height;
                }
            }
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

            var containers = parent.Controls.OfType<Panel>().DefaultIfEmpty().ToArray();

            for (int i = 1; i < containers.Length; i++ )
                containers[i].Left = containers[i - 1].Right;


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
