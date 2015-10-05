using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    class RibbonDropDown : SelectablePanel, IRibbonComponent
    {
        const string NEW_DROPDOWN_PREFIX = "New Dropdown";

        private TextBox _label = new TextBox();
        private ComboBox _cmb = new ComboBox();
        private Point _startPt = new Point(int.MaxValue, int.MaxValue);



        public string Descrizione { get; set; }
        public string ScreenTip { get; set; }
        //public string Nome { get; set; }
        public string Label { get { return _label.Text; } set { } }
        public int Slot { get { return 2; } }

        public RibbonDropDown(Control ribbon)
        {
            int prog = Utility.FindLastOfItsKind(ribbon, NEW_DROPDOWN_PREFIX, typeof(RibbonDropDown)) + 1;
            
            Name = NEW_DROPDOWN_PREFIX.Replace(" ", "") + prog;
            //_label.Font = ribbon.Font;
            _label.Text = NEW_DROPDOWN_PREFIX + " " + prog;
            _label.TextAlign = HorizontalAlignment.Left;
            _label.Click += SelectAllText;
            _label.Leave += CheckTextChanged;
            _label.MouseMove += ControlMouseMove;
            _cmb.MouseMove += ControlMouseMove;
            _label.MouseLeave += ControlMouseLeave;
            _cmb.MouseLeave += ControlMouseLeave;
            //Click += LeaveLabel;
            
            this.Font = ribbon.Font;

            //_cmb.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            //_txtLabel.Dock = DockStyle.Top;
            this.Padding = new Padding(4, (33 - _label.Height) / 2, 4, 4);
            this.Controls.Add(_cmb);
            this.Controls.Add(_label);
            _label.Top = Padding.Top;
            _label.Left = Padding.Left;
            //_label.Dock = DockStyle.Top;
            //_label.AutoSize = true;
            _label.BorderStyle = BorderStyle.FixedSingle;
            _label.BackColor = BackColor;
            _cmb.Top = _label.Bottom + 10;
            _cmb.Left = Padding.Left;
            //_cmb.FlatStyle = FlatStyle.Flat;

            Height = 66;
            _cmb.Width = 40;
            SetWidth();
        }

        private void SetWidth()
        {
            //SizeF s = Utility.MeasureTextSize(_label);
            int width = Math.Max(_label.GetPreferredSize(_label.Size).Width, _cmb.Width);
            _label.Width = width;
            this.Width = width + 2 * Padding.Left;
        }

        private void SelectAllText(object sender, EventArgs e)
        {
            _label.SelectAll();
        }

        private void CheckTextChanged(object sender, EventArgs e)
        {
            if (Name != _label.Text.Replace(" ", ""))
            {
                Name = _label.Text.Replace(" ", "");
                SetWidth();
                //_label.SelectAll();
                _label.SelectionStart = 0;
            }
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

        protected void ControlMouseMove(object sender, MouseEventArgs e)
        {
            base.OnMouseEnter(e);
            BackColor = Color.FromKnownColor(KnownColor.ControlDark);
            _label.BackColor = BackColor;
        }
        protected override void OnMouseMove(MouseEventArgs mevent)
        {
            if (mevent.Button == System.Windows.Forms.MouseButtons.Left && Math.Pow(mevent.Location.X - _startPt.X, 2) + Math.Pow(mevent.Location.Y - _startPt.Y, 2) > Math.Pow(SystemInformation.DragSize.Height, 2))
                DoDragDrop(this, DragDropEffects.Move);

            ControlMouseMove(this, mevent);
        }
        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            ControlMouseLeave(this, e);
        }
        protected override void OnMouseDown(MouseEventArgs mevent)
        {
            _startPt = mevent.Location;
            Select();
            if (mevent.Clicks == 2)
                OnDoubleClick(mevent);

            //base.OnMouseMove(mevent);
        }
        private void ControlMouseLeave(object sender, EventArgs e)
        {
            BackColor = Color.FromKnownColor(KnownColor.Control);
            _label.BackColor = BackColor;
        }
    }
}
