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
            _ws.Shapes.Item("lbTitolo").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(157, 58, 0));
            _ws.Shapes.Item("lbTitolo").Line.ForeColor.Brightness = 0.1019f;
            _ws.Shapes.Item("lbTitolo").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(198, 81, 12));
            _ws.Shapes.Item("lbTitolo").Fill.ForeColor.Brightness = 0.1792f;

            _ws.Shapes.Item("sfondo").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(157, 58, 0));
            _ws.Shapes.Item("sfondo").Line.ForeColor.Brightness = 0.1019f;
            _ws.Shapes.Item("sfondo").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 172, 123));
            _ws.Shapes.Item("sfondo").Fill.ForeColor.Brightness = 0.522f;

            _ws.Shapes.Item("lbDataInizio").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(198, 81, 12));
            _ws.Shapes.Item("lbDataInizio").Fill.ForeColor.Brightness = 0.1792f;
            _ws.Shapes.Item("lbDataFine").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(198, 81, 12));
            _ws.Shapes.Item("lbDataFine").Fill.ForeColor.Brightness = 0.1792f;


            //nascondi quelli non utilizzati
            _ws.Shapes.Item("lbImpianti").Visible = Office.MsoTriState.msoFalse;
            _ws.Shapes.Item("lbElsag").Visible = Office.MsoTriState.msoFalse;

            //sposto i due label sotto
            _ws.Shapes.Item("lbModifica").Top = _ws.Shapes.Item("lbImpianti").Top;
            _ws.Shapes.Item("lbTest").Top = _ws.Shapes.Item("lbElsag").Top;

            //ridimensiono lo sfondo
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoFalse;
            _ws.Shapes.Item("sfondo").Height = (float)(12.5 * _ws.Rows[5].Height);
            _ws.Shapes.Item("sfondo").LockAspectRatio = Office.MsoTriState.msoTrue;
        }
        protected override void Clear()
        {
            base.Clear();
            _ws.Rows[6].EntireRow.Hidden = true;
        }
    }
}
