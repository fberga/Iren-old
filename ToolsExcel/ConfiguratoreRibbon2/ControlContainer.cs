using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ConfiguratoreRibbon2
{
    class ControlContainer : Panel
    {
        public int FreeSlot { get; private set; }

        public ControlContainer()
        {
            BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            FreeSlot = 3;
        }

        protected override void OnControlAdded(ControlEventArgs e)
        {
            base.OnControlAdded(e);

            if (e.Control.GetType() == typeof(RibbonButton))
            {
                int dim = ((RibbonButton)e.Control).Dimensione;

                if (dim == 1)
                    FreeSlot = 0;
                else if (dim == 0)
                    FreeSlot -= 1;
            }
        }

        protected override void OnDragEnter(DragEventArgs drgevent)
        {
            Control ctrl = drgevent.Data as Control;

            base.OnDragEnter(drgevent);

            if (ctrl.GetType() == typeof(Button))
            {
                int dim = ((RibbonButton)ctrl).Dimensione;

                if (dim == 1 && FreeSlot < 3)
                    drgevent.Effect = DragDropEffects.None;
                else if (dim == 0 && FreeSlot == 0)
                    drgevent.Effect = DragDropEffects.None;
                else
                    drgevent.Effect = DragDropEffects.Move;
            }
        }
    }
}
