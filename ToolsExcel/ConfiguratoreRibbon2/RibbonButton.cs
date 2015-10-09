using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public class RibbonButton : SelectableButton, INotifyPropertyChanged, IRibbonControl
    {
        const string NEW_BUTTON_PREFIX = "New Button";

        private Size largeBtnMinSize = new Size(50, 100);
        private Size smallBtnMaxSize = new Size(250, 33);

        private ImageList _imageListNormal = new ImageList();
        private ImageList _imageListSmall = new ImageList();

        private Rectangle _selectionRect = new Rectangle(new Point(0, 0), new Size(0, 0));

        private Point _startPt = new Point(int.MaxValue, int.MaxValue);

        public int IdTipologia { get { return ToggleButton ? 2 : 1; } }
        public int Slot { get { return Dimensione == 1 ? 3 : 1; } }

        private int _dimensione = 1;
        public int Dimensione {
            get
            {
                return _dimensione;
            }
            set 
            {
                _dimensione = value;
                if (_dimensione == 1)
                {
                    SetUpLargeButton();
                    SetLargeButtonDimension();
                }
                else if (_dimensione == 0)
                {
                    SetUpSmallButton();
                    SetSmallButtonDimension();
                }
            } 
        }
        public string Descrizione { get; set; }
        public string ScreenTip { get; set; }
        public bool ToggleButton { get; set; }
        public string Label { get { return Text; } set { Text = value; } }
        public string ImageName { get { return ImageKey; } }
        public int ID { get; private set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public RibbonButton(ImageList normal, ImageList small, string imageKey, int id)
        {
            _imageListNormal = normal;
            _imageListSmall = small;
            ImageKey = imageKey;
            ID = id;
        }
        public RibbonButton(ImageList normal, ImageList small)
        {
            _imageListNormal = normal;
            _imageListSmall = small;

            SetUpLargeButton();
            Dimensione = 1;

            using (SelettoreImmagini chooseImageDialog = new SelettoreImmagini(_imageListNormal))
            {
                if (chooseImageDialog.ShowDialog() == DialogResult.OK)
                    ImageKey = chooseImageDialog.ResourceName;
            }
        }
        public RibbonButton(Control ribbon, ImageList normal, ImageList small)
            : this(normal, small)
        {
            int prog = Utility.FindLastOfItsKind(ribbon, NEW_BUTTON_PREFIX, typeof(RibbonButton)) + 1;
            Name = NEW_BUTTON_PREFIX.Replace(" ","") + prog;
            Text = NEW_BUTTON_PREFIX + " " + prog;
            this.Font = ribbon.Font;
            
            SetLargeButtonDimension();
        }


        public void SetUpLargeButton()
        {
            ImageList = _imageListNormal;
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
            ImageList = _imageListSmall;
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

        protected override void OnSizeChanged(EventArgs e)
        {
            if (Parent != null)
            {
                ControlContainer parent = Parent as ControlContainer;
                parent.SetContainerWidth();
            }
            base.OnSizeChanged(e);
        }

        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        protected override void OnDoubleClick(EventArgs e)
        {
            if (ID == 0)
            {
                int dim = Dimensione;

                using (ConfiguratoreTasto cfg = new ConfiguratoreTasto(this, _imageListNormal))
                {
                    cfg.ShowDialog();

                    if (dim != Dimensione)
                    {
                        OnPropertyChanged("Dimensione");
                    }
                }
            }
            else
            {
                RibbonGroup grp = Parent.Parent as RibbonGroup;

                AssegnaFunzioni afForm = new AssegnaFunzioni(this, grp, 1, 62);
                afForm.Show();
            }
            base.OnDoubleClick(e);
        }
        protected override void OnMouseDown(MouseEventArgs mevent)
        {
            _startPt = mevent.Location;
            if (mevent.Clicks == 2)
                OnDoubleClick(mevent);

            //base.OnMouseMove(mevent);
        }
        protected override void OnMouseMove(MouseEventArgs mevent)
        {
            if (mevent.Button == System.Windows.Forms.MouseButtons.Left && Math.Pow(mevent.Location.X - _startPt.X, 2) + Math.Pow(mevent.Location.Y - _startPt.Y, 2) > Math.Pow(SystemInformation.DragSize.Height, 2))
            {
                //rettangolo di spostamento
                //ControlPaint.DrawReversibleFrame(_selectionRect, this.BackColor, FrameStyle.Thick);
                //Point centerPoint = PointToScreen(new Point(mevent.X - DisplayRectangle.Width / 2, mevent.Y - DisplayRectangle.Height / 2));
                //_selectionRect = new Rectangle(centerPoint, this.DisplayRectangle.Size);
                //ControlPaint.DrawReversibleFrame(_selectionRect, this.BackColor, FrameStyle.Thick);


                DoDragDrop(this, DragDropEffects.Move);
            }

            //base.OnMouseMove(mevent);
        }
        protected override void OnMouseUp(MouseEventArgs mevent)
        {
            //if (mevent.Button == System.Windows.Forms.MouseButtons.Left)
            //{
            //    //ControlPaint.DrawReversibleFrame(_selectionRect, this.BackColor, FrameStyle.Thick);
            //    //_selectionRect = new Rectangle(new Point(0, 0), new Size(0, 0));
            //    OnClick(mevent);
            //}

            //base.OnMouseUp(mevent);
        }
        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            BackColor = Color.FromKnownColor(KnownColor.ControlDark);
        }
        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            BackColor = Color.FromKnownColor(KnownColor.Control);
        }
    }
}
