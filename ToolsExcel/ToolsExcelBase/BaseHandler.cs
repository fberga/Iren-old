using Iren.ToolsExcel.Core;
using System;
using System.Data;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.ToolsExcel.Base
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
                    string sheet = DefinedNames.GetSheetName(entita);
                    if (DefinedNames.IsDefined(sheet, DefinedNames.GetName(entita, "T", "DATA1")))
                    {
                        Target.Application.Worksheets[sheet].Activate();
                        Tuple<int, int> coordinate = definedNames[DefinedNames.GetName(entita, "T", "DATA1")][0];
                        Target.Application.Worksheets[sheet].Cells[coordinate.Item1, coordinate.Item2].Select();
                        Target.Application.ActiveWindow.SmallScroll(coordinate.Item1 - Target.Worksheet.Application.ActiveWindow.VisibleRange.Cells[1, 1].Row - 1);
                    }
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

        public static void StoreEdit(object Sh, Excel.Range Target)
        {
            DefinedNames nomiDefiniti = new DefinedNames(Target.Worksheet.Name);
            Sheet s = new Sheet(Target.Worksheet);
            Target.Worksheet.Unprotect(Simboli.pwd);
            s.AggiornaGrafici();

            object[,] values;
            if (Target.Value == null)   //caso in cui cancello il valore di una cella
            {
                values = new object[1, 1];
                values[0, 0] = null;
            }
            else if (Target.Value.GetType() != typeof(object[,]))   //caso in cui modifico il valore di una cella
            {
                values = new object[1,1];
                values[0,0] = Target.Value;
            }
            else    //caso in cui modifico un range di celle
            {
                values = new object[Target.Value.GetLength(0), Target.Value.GetLength(1)];
                Array.Copy(Target.Value, 1, values, 0, values.Length);
            }

            DataView modifiche = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.MODIFICA].DefaultView;

            for (int i = 0, rowLen = values.GetLength(0); i < rowLen; i++)
            {
                for (int j = 0, colLen = values.GetLength(1); j < colLen; j++)
                {
                    if (nomiDefiniti.SalvaDB(i + Target.Row, j + Target.Column))
                    {
                        string[] nomi = nomiDefiniti.Get(i + Target.Row, j + Target.Column);

                        string[] info = nomi[0].Split(Simboli.UNION[0]);
                        string data = CommonFunctions.GetDataFromSuffisso(info[2], info[3]);

                        modifiche.RowFilter = "SiglaEntita = '" + info[0] + "' AND SiglaInformazione = '" + info[1] + "' AND Data = '" + data + "'";
                        if (modifiche.Count == 0)
                            modifiche.Table.Rows.Add(info[0], info[1], data, values[i, j].ToString(), nomiDefiniti.AnnotaModifica(i + Target.Row, j + Target.Column), DataBase.IdApplicazione, DataBase.IdUtenteAttivo);
                        else
                            modifiche[0]["Valore"] = values[i, j];
                    }
                    if (nomiDefiniti.AnnotaModifica(i + Target.Row, j + Target.Column))
                    {
                        Excel.Range rng = Target.Worksheet.Cells[i + Target.Row, j + Target.Column];
                        rng.ClearComments();
                        rng.AddComment("Valore inserito manualmente").Visible = false;
                    }
                }
            }
            Target.Worksheet.Protect(Simboli.pwd);
        }

        public static void ChangeModificaDati(bool modifica)
        {
            Excel.Worksheet ws = CommonFunctions.WB.Sheets["Main"];
            ws.Unprotect(Simboli.pwd);
            ws.Shapes.Item("lbModifica").TextFrame.Characters().Text = "Modifica dati: " + (modifica ? "SI" : "NO");
            if (modifica) 
                ws.Shapes.Item("lbModifica").ShapeStyle = Office.MsoShapeStyleIndex.msoShapeStylePreset10;
            else
            {
                ws.Shapes.Item("lbModifica").BackgroundStyle = Office.MsoBackgroundStyleIndex.msoBackgroundStylePreset1;
                ws.Shapes.Item("lbModifica").TextFrame.Characters().Font.ColorIndex = 1;
                ws.Shapes.Item("lbModifica").Line.Weight = 0.75f;
                ws.Shapes.Item("lbModifica").Line.ForeColor.ObjectThemeColor = Office.MsoThemeColorIndex.msoThemeColorBackground1;
                ws.Shapes.Item("lbModifica").Line.ForeColor.TintAndShade = 0;
                ws.Shapes.Item("lbModifica").Line.ForeColor.Brightness = -0.25f;
                ws.Shapes.Item("lbModifica").Shadow.Visible = Office.MsoTriState.msoFalse;
            }

            ws.Protect(Simboli.pwd);
        }

        public static void ChangeAmbiente(string ambiente)
        {
            Excel.Worksheet ws = CommonFunctions.WB.Sheets["Main"];
            ws.Unprotect(Simboli.pwd);

            switch (ambiente)
            {
                case "Dev":
                    ws.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: DEVELOPMENT";
                    ws.Shapes.Item("lbTest").ShapeStyle = Office.MsoShapeStyleIndex.msoShapeStylePreset10;
                    break;
                case "Test":
                    ws.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: TEST";
                    ws.Shapes.Item("lbTest").ShapeStyle = Office.MsoShapeStyleIndex.msoShapeStylePreset14;
                    break;
                case "Produzione":
                    ws.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: PRODUZIONE";
                    ws.Shapes.Item("lbTest").ShapeStyle = Office.MsoShapeStyleIndex.msoShapeStylePreset11;
                    break;
            }

            ws.Protect(Simboli.pwd);
        }
    }
}
