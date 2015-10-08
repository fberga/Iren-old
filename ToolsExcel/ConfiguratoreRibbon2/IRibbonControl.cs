using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    interface IRibbonControl
    {
        int Slot { get; }
        string Descrizione { get; }
        string ImageName { get; }
        string ScreenTip { get; set; }
        string Label { get; set; }
        bool ToggleButton { get; }
        int Dimensione { get; }
        int IdTipologia { get; }
        int ID { get; }
    }
}
