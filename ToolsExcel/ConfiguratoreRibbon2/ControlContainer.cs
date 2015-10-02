using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    class ControlContainer : Panel
    {
        public int FreeSlot { get; private set; }
        public int CtrlCount { get; private set; }

        //private Dictionary<int, Control> _slots = new Dictionary<int, Control>();

        public ControlContainer()
        {
            BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            FreeSlot = 3;
            CtrlCount = 0;
            AllowDrop = true;
        }

        public void SetContainerWidth()
        {
            int width =
                Utility.GetAll(this)
                .Select(c => c.Width)
                .DefaultIfEmpty()
                .Max();

            Width = width == 0 ? 50 : width + 2;
        }

        protected override void OnControlAdded(ControlEventArgs e)
        {
            SetContainerWidth();
            CtrlCount += 1;

            FreeSlot -= ((IRibbonComponent)e.Control).Slot;

            if (e.Control.GetType() == typeof(RibbonButton))
            {
                RibbonButton btn = (RibbonButton)e.Control;
                btn.PropertyChanged += ButtonPropertyChanged;
            }

            base.OnControlAdded(e);
        }
        protected override void OnControlRemoved(ControlEventArgs e)
        {
            SetContainerWidth();
            CtrlCount -= 1;

            FreeSlot += ((IRibbonComponent)e.Control).Slot;

            if (e.Control.GetType() == typeof(RibbonButton))
            {
                RibbonButton btn = (RibbonButton)e.Control;
                btn.PropertyChanged += ButtonPropertyChanged;
            }

            CompactCtrls();

            base.OnControlRemoved(e);
        }

        private void CompactCtrls()
        {
            var ctrls = Controls;//Utility.GetAll(this).OrderBy(c => c.Top).ToList();
            if (ctrls.Count > 0)
            {
                ctrls[0].Top = 0;
                for (int i = 1; i < ctrls.Count; i++)
                    ctrls[i].Top = ctrls[i - 1].Bottom;
            }
        }

        private void ButtonPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            RibbonButton btn = sender as RibbonButton;
            if (btn.Parent == this && e.PropertyName == "Dimensione")
            {
                //non può essere diverso: o va a 1 e occupa tutto lo spazio, o va a 0 e occupa uno solo dei 3 slot
                FreeSlot = 3;
                FreeSlot -= btn.Slot;
            }
        }

        protected override void OnDragEnter(DragEventArgs drgevent)
        {
            Control ctrl = drgevent.Data.GetData(drgevent.Data.GetFormats()[0]) as Control;
            if (ctrl.Parent != this)
            {
                int slot = ((IRibbonComponent)ctrl).Slot;

                if (slot == 3 && FreeSlot < 3)
                    drgevent.Effect = DragDropEffects.None;
                else if (slot == 1 && FreeSlot == 0)
                    drgevent.Effect = DragDropEffects.None;
                else
                    drgevent.Effect = DragDropEffects.Move;
            }

            base.OnDragEnter(drgevent);
        }
        protected override void OnDragOver(DragEventArgs drgevent)
        {
            Control ctrl = drgevent.Data.GetData(drgevent.Data.GetFormats()[0]) as Control;
            
            //DisplaceObjects(drgevent, ctrl.Height);

            base.OnDragOver(drgevent);
        }

        //private void DisplaceObjects(DragEventArgs drgevent, int height)
        //{
        //    //trova a quale oggetto si sovrappone e sposta lui e tutti quelli sotto
        //    Point ptOver = PointToClient(new Point(drgevent.X, drgevent.Y));
        //    Control overlap = GetChildAtPoint(ptOver);
        //    if (overlap != null && overlap != drgevent.Data.GetData(drgevent.Data.GetFormats()[0]) as Control)
        //    {
        //        //trovo a quale slot appartiene
        //        var slot = _slots
        //                    .Where(kv => kv.Value.Equals(overlap))
        //                    .Select(kv => kv.Key).First();

        //        //metto sopra
        //        if (ptOver.Y - overlap.Top > overlap.Height / 2.0d)
        //        {
        //            for (int i = slot; i <= 3; i++)
        //            {
        //                if(_slots.ContainsKey(i))
        //                    _slots[i].Top = (i) * _slots[i].Height;
        //            }
        //        }
        //        //metto sotto
        //        else
        //        {
        //            for (int i = slot; i <= 3; i++)
        //            {
        //                if (_slots.ContainsKey(i))
        //                    _slots[i].Top = (i - 1) * _slots[i].Height;
        //            }
        //        }
                


        //        var ctrls =
        //        Utility.GetAll(this)
        //        .Where(c => c.Top >= overlap.Top);

        //        foreach (Control c in ctrls)
        //        {
        //            c.Top += height;
        //        }
        //    }
            
        //}

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (Parent != null)
            {
                Utility.UpdateGroupDimension(Parent);
            }
        }

        protected override void OnDragDrop(DragEventArgs drgevent)
        {            
            Control ctrl = drgevent.Data.GetData(drgevent.Data.GetFormats()[0]) as Control;

            int top = 0;
            if (ctrl != null)
            {
                
                int slot = ((IRibbonComponent)ctrl).Slot;
                if (slot < 3)
                {
                    top =
                        Utility.GetAll(this, typeof(IRibbonComponent))
                        .Select(b => b.Bottom)
                        .DefaultIfEmpty()
                        .Max();
                }

                Controls.Add(ctrl);
                ctrl.Left = 0;
                ctrl.Top = top;

            }
            base.OnDragDrop(drgevent);
        }
    }
}
