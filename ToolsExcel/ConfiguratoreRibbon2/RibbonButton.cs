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
        public const string NEW_BUTTON_PREFIX = "New Button";

        private Size largeBtnMinSize = new Size(50, 100);
        private Size smallBtnMaxSize = new Size(250, 33);

        private Point _startPt = new Point(int.MaxValue, int.MaxValue);

        public int IdTipologia { get { return ToggleButton ? 2 : 1; } }
        public int Slot { get { return Dimension == 1 ? 3 : 1; } }

        private int _dimensione = 1;
        public int Dimension {
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
        public string Description { get; set; }
        public string ScreenTip { get; set; }
        public bool ToggleButton { get; set; }
        public int IdControllo { get; private set; }
        public List<int> Functions { get; set; }
        //public bool Enabled { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        public RibbonButton(string imageKey, int id)
        {
            ImageKey = imageKey;
            IdControllo = id;
            Font = Utility.StdFont;
            Functions = new List<int>();
        }
        public RibbonButton(Control ribbon)
        {
            Font = Utility.StdFont;
            Functions = new List<int>();
            
            SetUpLargeButton();
            Dimension = 1;

            using (ConfiguratoreTasto configuraTasto = new ConfiguratoreTasto(this, ribbon))
            {
                if (configuraTasto.ShowDialog() != DialogResult.OK)
                {
                    this.Dispose();
                    return;
                }
            }

            SetLargeButtonDimension();
        }
        //public RibbonButton(Control ribbon)
        //    : this()
        //{
        //    int prog = Utility.FindLastOfItsKind(ribbon, NEW_BUTTON_PREFIX, typeof(RibbonButton)) + 1;
        //    Name = NEW_BUTTON_PREFIX.Replace(" ","") + prog;
        //    Text = NEW_BUTTON_PREFIX + " " + prog;
            
        //    SetLargeButtonDimension();
        //}

        public void SetUpLargeButton()
        {
            ImageList = Utility.ImageListNormal;
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
            ImageList = Utility.ImageListSmall;
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
            if (IdControllo == 0)
            {
                int dim = Dimension;

                using (ConfiguratoreTasto cfg = new ConfiguratoreTasto(this))
                {
                    cfg.ShowDialog();

                    if (dim != Dimension)
                    {
                        OnPropertyChanged("Dimensione");
                    }
                }
            }
            else
            {
                RibbonGroup grp = Parent.Parent as RibbonGroup;

                using (AssegnaFunzioni afForm = new AssegnaFunzioni(this, grp, 1, 62))
                {
                    if (afForm.ShowDialog() == DialogResult.OK)
                    {

                    }
                }
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

        protected override void Dispose(bool disposing)
        {            
            if (ConfiguratoreRibbon.ControlliUtilizzati.Contains(IdControllo))
                ConfiguratoreRibbon.GruppoControlloCancellati.Add(ConfiguratoreRibbon.GruppoControlloUtilizzati[ConfiguratoreRibbon.ControlliUtilizzati.IndexOf(IdControllo)]);

            base.Dispose(disposing);
        }
    }
}
