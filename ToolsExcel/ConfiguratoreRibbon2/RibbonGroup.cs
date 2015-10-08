using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    class RibbonGroup : SelectablePanel
    {
        const string NEW_GROUP_PREFIX = "New Group";

        TextBox _label = new TextBox();
        public string Label { get { return _label.Text; } set { _label.Text = value; } }

        public int ID { get; private set; }
        
        public RibbonGroup() 
            : base()
        {
            //Dock = DockStyle.Left;
            Padding = new Padding(4, 4, 4, 4);

            Controls.Add(_label);
            _label.Dock = DockStyle.Bottom;
            _label.AutoSize = false;
            _label.TextAlign = HorizontalAlignment.Center;
            _label.Click += SelectAllText;
            _label.Leave += CheckTextChanged;

            _label.BorderStyle = BorderStyle.None;
        }

        private void SelectAllText(object sender, EventArgs e)
        {
            _label.SelectAll();
        }

        private void CheckTextChanged(object sender, EventArgs e)
        {
            if(_label.Name != _label.Text.Replace(" ", ""))
            {
                _label.Name = _label.Text.Replace(" ", "");
                Utility.UpdateGroupDimension(this);
            }
        }

        public RibbonGroup(Control ribbon)
            : this()
        {
            BackColor = ControlPaint.LightLight(ribbon.BackColor);

            int prog = Utility.FindLastOfItsKind(ribbon, NEW_GROUP_PREFIX, typeof(TextBox)) + 1;

            Label = NEW_GROUP_PREFIX + " " + prog;

            Top = ribbon.Padding.Top;
            Width = (int)(Utility.MeasureTextSize(_label).Width + 20);
            Height = ribbon.Height - ribbon.Padding.Top - ribbon.Padding.Bottom;
            _label.BackColor = ControlPaint.LightLight(ribbon.BackColor);
        }
        public RibbonGroup(Control ribbon, int id)
            : this(ribbon)
        {
            ID = id;
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
            var rc = this.ClientRectangle;
            ControlPaint.DrawBorder3D(pe.Graphics, rc, Border3DStyle.Etched, Border3DSide.Right);
        }
        protected override void OnDoubleClick(EventArgs e)
        {
            base.OnDoubleClick(e);

            _label.Focus();
            _label.SelectAll();
        }
        protected override void OnControlAdded(ControlEventArgs e)
        {
            base.OnControlAdded(e);
            Utility.UpdateGroupDimension(this);
        }
        protected override void OnControlRemoved(ControlEventArgs e)
        {
            base.OnControlRemoved(e);
            CompactCtrls();
            Utility.UpdateGroupDimension(this);
        }
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);

            if(Parent != null)
                Utility.GroupsDisplacement(Parent);
        }

        private void CompactCtrls()
        {
            var ctrls = Controls
                .OfType<ControlContainer>()
                .OrderBy(c => c.Left)
                .ToList();

            if (ctrls.Count > 0)
            {
                ctrls[0].Left = Padding.Left;
                for (int i = 1; i < ctrls.Count; i++)
                    ctrls[i].Left = ctrls[i - 1].Right;
            }
        }
    }
}
