using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Data;
using System.Globalization;

namespace Iren.FrontOffice.Base
{
    public class Style
    {
        public static void SetAllBorders(Excel.Style s, int colorIndex, Excel.XlBorderWeight weight)
        {
            s.Borders.ColorIndex = 1;
            s.Borders.Weight = weight;
            s.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            s.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }

        public static void StdStyles(Workbook wb)
        {
            Excel.Style style;
            try
            {
                style = wb.Styles["gotoBarStyle"];
            }
            catch
            {
                style = wb.Styles.Add("gotoBarStyle");
                style.Font.Bold = false;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 15;


                style = wb.Styles.Add("navBarStyle");
                style.Font.Bold = true;
                style.Font.Size = 7;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 2;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);


                style = wb.Styles.Add("titleBarStyle");
                style.Font.Bold = true;
                style.Font.Size = 16;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 37;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);


                style = wb.Styles.Add("dateBarStyle");
                style.Font.Bold = true;
                style.Font.Size = 10;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.NumberFormat = "dddd d mmmm yyyy";
                style.Interior.ColorIndex = 15;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);


                style = wb.Styles.Add("chartsBarStyle");
                style.Font.Size = 10;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.NumberFormat = "dddd d mmmm yyyy";
                style.Interior.ColorIndex = 2;
                style.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;


                style = wb.Styles.Add("allDatiStyle");
                style.Font.Size = 10;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.NumberFormat = "#,###.000;-#,###.000;-";
                style.Interior.ColorIndex = 2;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);


                style = wb.Styles.Add("titoloVertStyle");
                style.Font.Bold = true;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 2;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);
            }
        }

        public static void RangeStyle(Excel.Range rng, string style)
        {
            MatchCollection paramsList = Regex.Matches(style, @"\w+[=:]((\[.*\])|([^;:=]+))");
            foreach (Match par in paramsList)
            {
                string[] keyVal;
                if (Regex.IsMatch(par.Value, @"\[.*\]"))
                    keyVal = Regex.Split(par.Value, @"[:=](?=\[.*\])");
                else
                    keyVal = Regex.Split(par.Value, @"[:=]");

                if (keyVal.Length != 2)
                    throw new InvalidExpressionException("The provided list of parameters is invalid.");

                switch (keyVal[0].ToLowerInvariant())
                {
                    case "merge":
                        rng.MergeCells = Regex.IsMatch(keyVal[1], "true|1", RegexOptions.IgnoreCase);
                        break;
                    case "bold":
                        rng.Font.Bold = Regex.IsMatch(keyVal[1], "true|1", RegexOptions.IgnoreCase);
                        break;
                    case "fontsize":
                        double size;
                        if (!double.TryParse(keyVal[1], out size))
                            size = 10.0;
                        rng.Font.Size = size;
                        break;
                    case "align":
                        string align = "xlHAlign" + Regex.Replace(keyVal[1], @"Center|Across|Selection|Distributed|Fill|General|Justify|Left|Right", delegate(Match m)
                        {
                            string v = m.ToString();
                            return char.ToUpper(v[0]) + v.Substring(1);
                        }, RegexOptions.IgnoreCase);

                        rng.HorizontalAlignment = (Excel.XlHAlign)Enum.Parse(typeof(Excel.XlHAlign), align);
                        rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        break;
                    case "numberformat":
                        rng.NumberFormat = keyVal[1];
                        break;
                    case "forecolor":
                        rng.Font.ColorIndex = int.Parse(keyVal[1]);
                        break;
                    case "backcolor":
                        rng.Interior.ColorIndex = int.Parse(keyVal[1]);
                        break;
                    case "backpattern":
                        string backPattern = "xlPattern" + Regex.Replace(keyVal[1], "Vertical|Up|None|Horizontal|Gray|Down|Automatic|Solid|Checker|Semi|Light|Grid|Criss|Cross|Linear|Gradient|Rectangular", delegate(Match m)
                        {
                            string v = m.ToString();
                            return char.ToUpper(v[0]) + v.Substring(1);
                        }, RegexOptions.IgnoreCase);

                        rng.Interior.Pattern = (Excel.XlPattern)Enum.Parse(typeof(Excel.XlPattern), backPattern);
                        break;
                    case "borders":
                        MatchCollection borders = Regex.Matches(keyVal[1], @"(Top|Left|Bottom|Right|InsideH|InsideV)([:=]\w*)?", RegexOptions.IgnoreCase);
                        rng.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                        foreach (Match border in borders)
                        {
                            string[] b = Regex.Split(border.Value, @"[:=]\s*");

                            Excel.XlBordersIndex index = Excel.XlBordersIndex.xlEdgeTop;
                            Excel.XlBorderWeight weight = Excel.XlBorderWeight.xlThin;
                            switch (b[0].ToLowerInvariant())
                            {
                                case "top":
                                    index = Excel.XlBordersIndex.xlEdgeTop;
                                    break;
                                case "left":
                                    index = Excel.XlBordersIndex.xlEdgeLeft;
                                    break;
                                case "bottom":
                                    index = Excel.XlBordersIndex.xlEdgeBottom;
                                    break;
                                case "right":
                                    index = Excel.XlBordersIndex.xlEdgeRight;
                                    break;
                                case "insideh":
                                    index = Excel.XlBordersIndex.xlInsideHorizontal;
                                    break;
                                case "insidev":
                                    index = Excel.XlBordersIndex.xlInsideVertical;
                                    break;
                            }
                            if (b.Length == 2)
                            {
                                switch (b[1].ToLowerInvariant())
                                {
                                    case "thick":
                                        weight = Excel.XlBorderWeight.xlThick;
                                        break;
                                    case "thin":
                                        weight = Excel.XlBorderWeight.xlThin;
                                        break;
                                    case "medium":
                                        weight = Excel.XlBorderWeight.xlMedium;
                                        break;
                                    case "hairline":
                                        weight = Excel.XlBorderWeight.xlHairline;
                                        break;
                                }
                            }
                            rng.Borders[index].LineStyle = Excel.XlLineStyle.xlContinuous;
                            rng.Borders[index].Weight = weight;
                        }
                        break;
                    case "orientation":
                        string orientation = "xl" + Regex.Replace(keyVal[1], "Downward|Horizontal|Upward|Vertical", delegate(Match m)
                        {
                            string v = m.ToString();
                            return char.ToUpper(v[0]) + v.Substring(1);
                        }, RegexOptions.IgnoreCase);

                        rng.Orientation = (Excel.XlOrientation)Enum.Parse(typeof(Excel.XlOrientation), orientation);
                        break;
                    case "visible":
                        rng.EntireRow.Hidden = Regex.IsMatch(keyVal[1], "false|0", RegexOptions.IgnoreCase);
                        break;
                }
            }
        }
    }
}
