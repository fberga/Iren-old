using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    interface IRibbonComponent
    {
        int Slot { get; }
        string Descrizione { get; set; }
        string ScreenTip { get; set; }
        string Label { get; set; }
        string Nome { get; set; }
    }
}
