using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Iren.FrontOffice.Core;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.FrontOffice.Base
{
    public class BaseHandler
    {
        public static void GotoClick(object Sh, Excel.Range Target)
        {
            DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name);

            string[] names = definedNames.Get(Target.Row, Target.Column) ?? new string[0];
            bool isGOTO = false;
            int i = 0;
            while(!isGOTO && i < names.Length ) 
            {
                isGOTO = Regex.IsMatch(names[i], "GOTO");
                i++;
            }

            if (isGOTO)
            {
                string entita = Regex.Replace(names[i - 1], "(RIEPILOGO" + Simboli.UNION + "|" + Simboli.UNION + "GOTO)", "");

                if (Target.Worksheet.Name == "Main")
                {
                    string sheet = definedNames.GetSheetName(entita);
                    Target.Application.Worksheets[sheet].Activate();

                    Tuple<int,int> coordinate = definedNames[entita + Simboli.UNION + "T" + Simboli.UNION + "DATA1"][0];
                    Target.Application.Worksheets[sheet].Cells[coordinate.Item1, coordinate.Item2].Select();
                    Target.Application.ActiveWindow.SmallScroll(coordinate.Item1 - Target.Worksheet.Application.ActiveWindow.VisibleRange.Cells[1, 1].Row - 1);
                }
                else
                {
                    Tuple<int, int> coordinate = definedNames[entita + Simboli.UNION + "T" + Simboli.UNION + "DATA1"][0];
                    Excel.Range rng = Target.Worksheet.Cells[coordinate.Item1, coordinate.Item2];
                    rng.Select();
                    Target.Worksheet.Application.ActiveWindow.SmallScroll(rng.Row - Target.Worksheet.Application.ActiveWindow.VisibleRange.Cells[1, 1].Row - 1);
                }

            }
        }
    }
}
