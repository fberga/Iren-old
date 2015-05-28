using Iren.ToolsExcel.Utility;
using Microsoft.Office.Tools.Excel;
using System;
using System.Data;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    public class Style : Utility.Workbook
    {
        public static void SetAllBorders(Excel.Style s, int colorIndex, Excel.XlBorderWeight weight)
        {
            s.Borders.ColorIndex = colorIndex;
            s.Borders.Weight = weight;
            s.Borders[Excel.XlBordersIndex.xlDiagonalDown].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            s.Borders[Excel.XlBordersIndex.xlDiagonalUp].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }

        public static void StdStyles()
        {
            Microsoft.Office.Tools.Excel.Workbook wb = WB;
            Excel.Style style;
            try
            {
                style = wb.Styles["gotoBarStyle"];
            }
            catch
            {
                style = wb.Styles.Add("gotoBarStyle");
                style.Font.Bold = false;
                style.Font.Name = "Verdana";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 15;

                style = wb.Styles.Add("navBarStyleVertical");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 8;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 2;
                style.NumberFormat = "ddd d";
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);

                style = wb.Styles.Add("navBarStyleHorizontal");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 8;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 2;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);


                style = wb.Styles.Add("titleBarStyle");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 16;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 37;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);

                style = wb.Styles.Add("dateBarStyle");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 12;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.NumberFormat = "dddd d mmmm yyyy";
                style.Interior.ColorIndex = 15;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);

                style = wb.Styles.Add("chartsBarStyle");
                style.Font.Size = 10;
                style.Font.Name = "Verdana";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.NumberFormat = "dddd d mmmm yyyy";
                style.Interior.ColorIndex = 2;
                style.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;

                style = wb.Styles.Add("allDatiStyle");
                style.Font.Size = 10;
                style.Font.Name = "Verdana";
                style.NumberFormat = "#,##0.0;-#,##0.0;-";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 2;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);                

                style = wb.Styles.Add("titoloVertStyle");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 2;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);

                style = wb.Styles.Add("recapTitleBarStyle");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Interior.ColorIndex = 37;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlMedium);

                style = wb.Styles.Add("recapEntityBarStyle");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 9;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                style.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;
                //SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);

                style = wb.Styles.Add("recapAllDatiStyle");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 9;
                style.Interior.ColorIndex = 2;
                style.Interior.Pattern = Excel.XlPattern.xlPatternCrissCross;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);

                style = wb.Styles.Add("recapCategoryTitle");
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 9;
                style.Interior.ColorIndex = 44;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                style = wb.Styles.Add("recapOKCell");
                style.Font.ColorIndex = 1;
                style.Font.Bold = true;
                style.Font.Name = "Verdana";
                style.Font.Size = 9;
                style.Interior.ColorIndex = 4;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                style = wb.Styles.Add("recapNPCell");
                style.Font.ColorIndex = 3;
                style.Font.Bold = false;
                style.Font.Name = "Verdana";
                style.Font.Size = 7;
                style.Interior.ColorIndex = 2;
                style.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                style = wb.Styles.Add("Adjustable");
                style.Font.Color = System.Drawing.Color.Coral;
                style.Font.Size = 10;
                style.Interior.ColorIndex = 35;
                style.NumberFormat = "#,##0.0;-#,##0.0;-";
                style.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                SetAllBorders(style, 1, Excel.XlBorderWeight.xlThin);
            }
        }

        public static void RangeStyle(Excel.Range rng, object fontName = null, object style = null, object merge = null, object bold = null, object fontSize = null, object align = null, object numberFormat = null, object foreColor = null, object backColor = null, object pattern = null, object borders = null, object orientation = null, object visible = null)
        {
            if(fontName != null)
                rng.Font.Name = (string)fontName;
            
            if(bold != null)
                rng.Font.Bold = (bool)bold;
            
            if(fontSize != null)
                rng.Font.Size = (int)fontSize;
            
            if(foreColor != null)
                rng.Font.ColorIndex = (int)foreColor;
            
            if(style != null)
                rng.Style = (string)style;
            
            if(merge != null)
                rng.MergeCells = (bool)merge;

            if (align != null)
            {
                rng.HorizontalAlignment = (Excel.XlHAlign)align;
                rng.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            }

            if(numberFormat != null)
                rng.NumberFormat = (string)numberFormat;

            if(backColor != null)
                rng.Interior.ColorIndex = (int)backColor;
            
            if(pattern != null)
                rng.Interior.Pattern = (Excel.XlPattern)pattern;

            if(orientation != null)
                rng.Orientation = (Excel.XlOrientation)orientation;

            if(visible != null)
                rng.EntireRow.Hidden = !(bool)visible;

            if (borders != null)
            {
                MatchCollection borderString = Regex.Matches((string)borders, @"(Top|Left|Bottom|Right|InsideH|InsideV)([:=]\w*)?", RegexOptions.IgnoreCase);
                foreach (Match border in borderString)
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
            }
        }
    }
}
