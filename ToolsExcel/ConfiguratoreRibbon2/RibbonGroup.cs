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
        public const string NEW_GROUP_PREFIX = "New Group";

        private TextBox Label { get; set; }
        public string Text { get { return Label.Text; } set { Label.Text = value; } }
        public string Name { get; set; }

        public int ID { get; private set; }
        
        public RibbonGroup() 
            : base()
        {
            //Dock = DockStyle.Left;
            this.Padding = new Padding(4, 4, 4, 4);

            this.Label = new TextBox();

            this.Controls.Add(this.Label);
            this.Label.Dock = DockStyle.Bottom;
            this.Label.AutoSize = false;
            this.Label.TextAlign = HorizontalAlignment.Center;
            this.Label.Click += SelectAllText;
            this.Label.Leave += CheckTextChanged;

            this.Label.BorderStyle = BorderStyle.None;
        }

        private void SelectAllText(object sender, EventArgs e)
        {
            this.Label.SelectAll();
        }

        private void CheckTextChanged(object sender, EventArgs e)
        {
            if (this.Label.Name != this.Text)
            {
                this.Label.Name = this.Text;
                Utility.UpdateGroupDimension(this);
            }
        }

        public RibbonGroup(Control ribbon)
            : this()
        {
            using (ConfiguraControllo cc = new ConfiguraControllo(ribbon, typeof(RibbonGroup)))
            {
                if (cc.ShowDialog() == DialogResult.OK)
                {
                    Name = cc.CtrlName;
                    Text = cc.CtrlText;
                }
                else
                {
                    Dispose();
                    return;
                }
            }

            BackColor = ControlPaint.LightLight(ribbon.BackColor);

            this.Top = ribbon.Padding.Top;
            this.Width = (int)(Utility.MeasureTextSize(this.Label).Width + 20);
            this.Height = ribbon.Height - ribbon.Padding.Top - ribbon.Padding.Bottom;
            this.Label.BackColor = ControlPaint.LightLight(ribbon.BackColor);
        }
        public RibbonGroup(Control ribbon, int id)
            : this(ribbon)
        {
            this.ID = id;
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
            var rc = this.ClientRectangle;
            ControlPaint.DrawBorder3D(pe.Graphics, rc, Border3DStyle.Etched, Border3DSide.Right);
        }
        //protected override void OnDoubleClick(EventArgs e)
        //{
        //    this.Label.Focus();
        //    this.Label.SelectAll();

        //    base.OnDoubleClick(e);
        //}
        protected override void OnControlAdded(ControlEventArgs e)
        {
            base.OnControlAdded(e);
            Utility.UpdateGroupDimension(this);
        }
        protected override void OnControlRemoved(ControlEventArgs e)
        {
            this.CompactCtrls();
            Utility.UpdateGroupDimension(this);

            base.OnControlRemoved(e);
        }
        protected override void OnSizeChanged(EventArgs e)
        {
            if(Parent != null)
                Utility.GroupsDisplacement(Parent);
            
            base.OnSizeChanged(e);
        }

        private void CompactCtrls()
        {
            var ctrls = this.Controls
                .OfType<ControlContainer>()
                .OrderBy(c => c.Left)
                .ToList();

            if (ctrls.Count > 0)
            {
                ctrls[0].Left = this.Padding.Left;
                for (int i = 1; i < ctrls.Count; i++)
                    ctrls[i].Left = ctrls[i - 1].Right;
            }
        }
    }
}