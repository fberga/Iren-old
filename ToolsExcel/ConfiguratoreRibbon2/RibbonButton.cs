using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ConfiguratoreRibbon2
{
    public class RibbonButton : Button
    {
        const string NEW_BUTTON_PREFIX = "New Button";

        private Size largeBtnMinSize = new Size(50, 100);
        private Size smallBtnMaxSize = new Size(250, 33);

        private ImageList imageListNormal = new ImageList();
        private ImageList imageListSmall = new ImageList();

        public int Dimensione { get; set; }
        public string Descrizione { get; set; }
        public string ScreenTip { get; set; }
        public bool ToggleButton { get; set; }
        public string Label { get { return Text; } set { Text = value; } }
        public string Nome { get { return Name; } set { Name = value; } }


        public RibbonButton(ImageList normal, ImageList small)
        {
            SetUpLargeButton();
            Dimensione = 1;

            using (SelettoreImmagini chooseImageDialog = new SelettoreImmagini(imageListNormal))
            {
                if (chooseImageDialog.ShowDialog() == DialogResult.OK)
                {
                    ImageKey = chooseImageDialog.FileName;
                    SetLargeButtonDimension();
                }
            }

            imageListNormal = normal;
            imageListSmall = small;
        }
        public RibbonButton(Control ribbon, ImageList normal, ImageList small)
            : this(normal, small)
        {
            int prog = Utility.FindLastOfItsKind(ribbon, NEW_BUTTON_PREFIX, typeof(Button)) + 1;
            Name = NEW_BUTTON_PREFIX.Replace(" ","") + prog;
            Text = NEW_BUTTON_PREFIX + " " + prog;
        }


        public void SetUpLargeButton()
        {
            ImageList = imageListNormal;
            MaximumSize = new Size(int.MaxValue, int.MaxValue);
            MinimumSize = largeBtnMinSize;
            ImageAlign = ContentAlignment.TopCenter;
            TextImageRelation = TextImageRelation.ImageAboveText;
            TextAlign = ContentAlignment.MiddleCenter;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
        }
        public void SetUpSmallButton()
        {
            ImageList = imageListSmall;
            MinimumSize = new Size(0, 0);
            MaximumSize = smallBtnMaxSize;
            ImageAlign = ContentAlignment.MiddleLeft;
            TextImageRelation = TextImageRelation.ImageBeforeText;
            TextAlign = ContentAlignment.MiddleLeft;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            AutoEllipsis = true;
        }

        public void SetLargeButtonDimension()
        {
            Width = Math.Min((int)(Utility.MeasureTextSize(this).Width + 15), 250);
            Height = MinimumSize.Height;
        }
        public void SetSmallButtonDimension()
        {
            Width = Math.Min((int)(Utility.MeasureTextSize(this).Width + 30), 250);
        }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);

            using (ConfiguratoreTasto cfg = new ConfiguratoreTasto(this, imageListNormal))
            {
                cfg.ShowDialog();

                if (Dimensione == 1)
                {
                    SetUpLargeButton();
                    SetLargeButtonDimension();
                }
                else if (Dimensione == 0)
                {
                    SetUpSmallButton();
                    SetSmallButtonDimension();
                }
            }
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);

            if (Parent != null)
            {
                Control parent = Parent;
                var maxBtnWidth =
                    (from btn in parent.Controls.OfType<Button>()
                     select btn.Width).DefaultIfEmpty().Max();
                var maxCmbWidth =
                    (from cmb in parent.Controls.OfType<ComboBox>()
                     select cmb.Width).DefaultIfEmpty().Max();
                var maxLabelWidth =
                    (from lbl in parent.Controls.OfType<Label>()
                     select (int)Utility.MeasureTextSize(lbl).Width).DefaultIfEmpty().Max();

                parent.Width = Enumerable.Max(new int[] { maxBtnWidth, maxCmbWidth, maxLabelWidth });
            }
        }
    }
}
