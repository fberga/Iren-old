using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.ToolsExcel
{
    class Riepilogo : Base.Riepilogo
    {
        public Riepilogo()
            : base()
        {

        }

        public Riepilogo(Excel.Worksheet ws)
            : base(ws)
        {

        }

        public override void InitLabels()
        {
            base.InitLabels();

            //coloro
            _ws.Shapes.Item("lbTitolo").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 42, 33));
            _ws.Shapes.Item("lbTitolo").Line.ForeColor.Brightness = 0.1114f;
            _ws.Shapes.Item("lbTitolo").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(15, 69, 58));
            _ws.Shapes.Item("lbTitolo").Fill.ForeColor.Brightness = 0.2024f;

            _ws.Shapes.Item("sfondo").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 42, 33));
            _ws.Shapes.Item("sfondo").Line.ForeColor.Brightness = 0.1114f;
            _ws.Shapes.Item("sfondo").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(94, 135, 127));
            _ws.Shapes.Item("sfondo").Fill.ForeColor.Brightness = 0.4778f;

            _ws.Shapes.Item("lbDataInizio").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(15, 69, 58));
            _ws.Shapes.Item("lbDataInizio").Fill.ForeColor.Brightness = 0.2024f;
            _ws.Shapes.Item("lbDataFine").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(15, 69, 58));
            _ws.Shapes.Item("lbDataFine").Fill.ForeColor.Brightness = 0.2024f;
        }
    }
}
