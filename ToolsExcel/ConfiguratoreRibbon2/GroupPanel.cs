﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace ConfiguratoreRibbon2
{
    class GroupPanel : SelectablePanel
    {
        const string NEW_GROUP_PREFIX = "New Group";

        TextBox _label = new TextBox();
        public string Label { get { return _label.Text; } set { _label.Text = value; } }

        public GroupPanel() 
            : base()
        {
            Dock = DockStyle.Left;
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

        public GroupPanel(Control ribbon)
            : this()
        {
            BackColor = ControlPaint.LightLight(ribbon.BackColor);

            int prog = Utility.FindLastOfItsKind(ribbon, NEW_GROUP_PREFIX, typeof(TextBox)) + 1;

            _label.Name = NEW_GROUP_PREFIX.Replace(" ", "") + prog;
            _label.Text = NEW_GROUP_PREFIX + " " + prog;

            Width = (int)(Utility.MeasureTextSize(_label).Width + 20);

            _label.BackColor = ControlPaint.LightLight(ribbon.BackColor);
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
            var rc = this.ClientRectangle;
            ControlPaint.DrawBorder3D(pe.Graphics, rc, Border3DStyle.Etched, Border3DSide.Right);
        }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
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
    }
}
