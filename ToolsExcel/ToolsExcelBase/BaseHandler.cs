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
                values = new object[1, 1];
                values[0, 0] = Target.Value;
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
                            modifiche.Table.Rows.Add(info[0], info[1], data, values[i, j].ToString(), nomiDefiniti.AnnotaModifica(i + Target.Row, j + Target.Column), CommonFunctions.DB.IdApplicazione, CommonFunctions.DB.IdUtenteAttivo);
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
            ws.Shapes.Item("lbModifica").Locked = false;
            ws.Shapes.Item("lbModifica").TextFrame.Characters().Text = "Modifica dati: " + (modifica ? "SI" : "NO");
            if (modifica) 
            {
                //giallo
                ws.Shapes.Item("lbModifica").Line.Weight = 2f;
                ws.Shapes.Item("lbModifica").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 204, 0));
                ws.Shapes.Item("lbModifica").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 102));
            }
            else
            {
                //bianco normale
                ws.Shapes.Item("lbModifica").Line.Weight = 0.75f;
                ws.Shapes.Item("lbModifica").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                ws.Shapes.Item("lbModifica").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                //ws.Shapes.Item("lbModifica").Line.ForeColor.Brightness = -0.25f;
            }
            ws.Shapes.Item("lbModifica").Locked = true;
        }
        public static void ChangeAmbiente(string ambiente)
        {
            Excel.Worksheet ws = CommonFunctions.WB.Sheets["Main"];
            ws.Shapes.Item("lbTest").Locked = false;
            switch (ambiente)
            {
                case "Dev":
                    ws.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: DEVELOPMENT";
                    //rosso
                    ws.Shapes.Item("lbTest").Line.Weight = 2f;
                    ws.Shapes.Item("lbTest").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(140, 56, 54));
                    ws.Shapes.Item("lbTest").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 80, 77));
                    break;
                case "Test":
                    ws.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: TEST";
                    //giallo
                    ws.Shapes.Item("lbTest").Line.Weight = 2f;
                    ws.Shapes.Item("lbTest").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 204, 0));
                    ws.Shapes.Item("lbTest").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 102));
                    break;
                case "Produzione":
                    ws.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: PRODUZIONE";
                    //bianco normale
                    ws.Shapes.Item("lbTest").Line.Weight = 0.75f;
                    ws.Shapes.Item("lbTest").BackgroundStyle = Office.MsoBackgroundStyleIndex.msoBackgroundStylePreset1;
                    ws.Shapes.Item("lbTest").Line.ForeColor.ObjectThemeColor = Office.MsoThemeColorIndex.msoThemeColorBackground1;
                    ws.Shapes.Item("lbTest").Line.ForeColor.Brightness = -0.25f;
                    break;
            }
            ws.Shapes.Item("lbTest").Locked = true;
        }
        public static void ChangeStatoDB(DataBase.NomiDB db, bool online)
        {
            string labelName = "";
            string labelText = "";
            switch (db)
            {
                case DataBase.NomiDB.SQLSERVER:
                    labelName = "lbSQLServer";
                    labelText = "Database SQL Server: ";
                    break;
                case DataBase.NomiDB.IMP:
                    labelName = "lbImpianti";
                    labelText = "Database Impianti: ";
                    break;
                case DataBase.NomiDB.ELSAG:
                    labelName = "lbElsag";
                    labelText = "Database Elsag: ";
                    break;
            }

            Excel.Worksheet ws = CommonFunctions.WB.Sheets["Main"];
            ws.Shapes.Item(labelName).TextFrame.Characters().Text = labelText + (online ? "OPERATIVO" : "FUORI SERVIZIO");
            if (online)
            {
                //bianco normale
                ws.Shapes.Item(labelName).Line.Weight = 0.75f;
                ws.Shapes.Item(labelName).Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                ws.Shapes.Item(labelName).Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                ws.Shapes.Item(labelName).Line.ForeColor.Brightness = -0.25f;
            }
            else
            {
                //rosso
                ws.Shapes.Item(labelName).Line.Weight = 2f;
                ws.Shapes.Item(labelName).Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(140, 56, 54));
                ws.Shapes.Item(labelName).Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 80, 77));
            }
        }
    }
}
