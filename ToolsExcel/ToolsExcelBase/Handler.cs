using System;
using System.Data;
using System.Linq;
using System.Resources;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Iren.ToolsExcel.Utility;
using System.Collections.Generic;
using System.IO;

namespace Iren.ToolsExcel.Base
{
    public class Handler
    {
        public static void GotoClick(object Sh, Excel.Range Target)
        {
            NewDefinedNames newDefinedNames = new NewDefinedNames(Target.Worksheet.Name);

            string address = newDefinedNames.GetGOTO(Target.Row, Target.Column);
            if (address != "")
            {

            }

            DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name);

            string[] names = definedNames.Get(Target.Row, Target.Column) ?? new string[0];

            bool isGOTO = false;
            int i = 0;
            while(!isGOTO && i < names.Length ) 
            {
                System.Windows.Forms.MessageBox.Show(names[i]);

                isGOTO = Regex.IsMatch(names[i], "GOTO");
                i++;
            }

            if (isGOTO)
            {
                string entita = Regex.Replace(names[i - 1], "(RIEPILOGO" + Simboli.UNION + "|" + Simboli.UNION + "GOTO)", "");
                GOTO(entita);
            }
        }

        /// <summary>
        /// Trova il foglio e la posizione dell'entita data in input. Se esiste, sposta la selezione al titolo.
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'UP su cui spostare la selezione</param>
        public static void GOTO(object siglaEntita)
        {
            string nomeFoglio = Workbook.WB.Application.ActiveSheet.Name;
            DefinedNames nomiDefiniti = new DefinedNames(nomeFoglio);
            string suffissoData = "";
            if (Struct.tipoVisualizzazione == "V")
            {
                string[] split = siglaEntita.ToString().Split(Simboli.UNION[0]);
                if (split.Length > 1)
                {
                    siglaEntita = split[0];
                    suffissoData = split[1];
                }
            }
            if (nomeFoglio == "Main" || !nomiDefiniti.IsDefined(siglaEntita.ToString()))
            {
                nomeFoglio = DefinedNames.GetSheetName(siglaEntita);
                if (nomeFoglio == null)
                    return;

                nomiDefiniti = new DefinedNames(nomeFoglio);
                if (!nomiDefiniti.IsDefined(DefinedNames.GetName(siglaEntita, "T", suffissoData == "" ? "DATA1" : suffissoData)))
                    return;

                Workbook.WB.Worksheets[nomeFoglio].Activate();
            }
            Tuple<int, int> coordinate = nomiDefiniti[DefinedNames.GetName(siglaEntita, "T", suffissoData == "" ? "DATA1" : suffissoData)][0];
            Workbook.WB.ActiveSheet.Cells[coordinate.Item1, coordinate.Item2].Select();
            Workbook.WB.Application.ActiveWindow.SmallScroll(coordinate.Item1 - Workbook.WB.Application.ActiveWindow.VisibleRange.Cells[1, 1].Row - 1);
        }

        public static void StoreEdit(object Sh, Excel.Range Target)
        {
            DefinedNames nomiDefiniti = new DefinedNames(Target.Worksheet.Name);
            Sheet s = new Sheet(Target.Worksheet);
            Target.Worksheet.Unprotect(Simboli.pwd);
            s.AggiornaGrafici();

            if (nomiDefiniti.SalvaDB(Target.Row, Target.Column))
            {
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

                DataView modifiche = DataBase.LocalDB.Tables[DataBase.Tab.MODIFICA].DefaultView;

                for (int i = 0, rowLen = values.GetLength(0); i < rowLen; i++)
                {
                    for (int j = 0, colLen = values.GetLength(1); j < colLen; j++)
                    {
                        if (values[i, j] != null)
                        {
                            if (nomiDefiniti.SalvaDB(i + Target.Row, j + Target.Column))
                            {
                                string[] nomi = nomiDefiniti.Get(i + Target.Row, j + Target.Column);

                                string[] info = nomi[0].Split(Simboli.UNION[0]);
                                string data = Utility.Date.GetDataFromSuffisso(info[2], info.Length == 4 ? info[3] : null);

                                modifiche.RowFilter = "SiglaEntita = '" + info[0] + "' AND SiglaInformazione = '" + info[1] + "' AND Data = '" + data + "'";
                                if (modifiche.Count == 0)
                                    modifiche.Table.Rows.Add(info[0], info[1], data, values[i, j].ToString(), nomiDefiniti.AnnotaModifica(i + Target.Row, j + Target.Column), DataBase.DB.IdApplicazione, DataBase.DB.IdUtenteAttivo);
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
                } 
            }
            Target.Worksheet.Protect(Simboli.pwd);
        }

        public static void ChangeModificaDati(bool modifica)
        {
            Excel.Worksheet ws = Workbook.WB.Sheets["Main"];
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
                ws.Shapes.Item("lbModifica").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 0, 0));
                ws.Shapes.Item("lbModifica").Line.ForeColor.Brightness = +0.75f;
            }
            ws.Shapes.Item("lbModifica").Locked = true;
        }
        public static void ChangeAmbiente(string ambiente)
        {
            Excel.Worksheet ws = Workbook.WB.Sheets["Main"];
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
                    ws.Shapes.Item("lbTest").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                    ws.Shapes.Item("lbTest").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 0, 0));
                    ws.Shapes.Item("lbTest").Line.ForeColor.Brightness = +0.75f;
                    break;
            }
            ws.Shapes.Item("lbTest").Locked = true;
        }
        public static void ChangeStatoDB(Core.DataBase.NomiDB db, bool online)
        {
            string labelName = "";
            string labelText = "";
            switch (db)
            {
                case Core.DataBase.NomiDB.SQLSERVER:
                    labelName = "lbSQLServer";
                    labelText = "Database SQL Server: ";
                    break;
                case Core.DataBase.NomiDB.IMP:
                    labelName = "lbImpianti";
                    labelText = "Database Impianti: ";
                    break;
                case Core.DataBase.NomiDB.ELSAG:
                    labelName = "lbElsag";
                    labelText = "Database Elsag: ";
                    break;
            }

            Excel.Worksheet ws = Workbook.WB.Sheets["Main"];
            var locked = ws.ProtectContents;
            if (locked)
                ws.Unprotect(Simboli.pwd);
            ws.Shapes.Item(labelName).TextFrame.Characters().Text = labelText + (online ? "OPERATIVO" : "FUORI SERVIZIO");
            if (online)
            {
                //bianco normale
                ws.Shapes.Item(labelName).Line.Weight = 0.75f;
                ws.Shapes.Item(labelName).Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                ws.Shapes.Item(labelName).Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 0, 0));
                ws.Shapes.Item(labelName).Line.ForeColor.Brightness = +0.75f;
            }
            else
            {
                //rosso
                ws.Shapes.Item(labelName).Line.Weight = 2f;
                ws.Shapes.Item(labelName).Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(140, 56, 54));
                ws.Shapes.Item(labelName).Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 80, 77));
            }

            if (locked)
                ws.Protect(Simboli.pwd);
        }

        public static void SwitchWorksheet(string name)
        {
            //TODO aprire gli altri file!!!!!!
            string path = "";
#if(DEBUG)
            path = "D:\\Repository\\Iren\\ToolsExcel\\"+name+"\\bin\\Debug\\"+name+".xlsm";
#else

#endif
            if (!IsWorkbookOpen(name))
            {
                Workbook.WB.Application.Workbooks.Open(path);
            }
            else
            {
                try
                {
                    Workbook.WB.Application.Workbooks[name + ".xlsm"].Activate();
                }
                catch(Exception)
                {
                    Workbook.WB.Application.Workbooks[name].Activate();
                }
            }
        }

        private static bool IsWorkbookOpen(string name)
        {
            try
            {
                Microsoft.Office.Interop.Excel._Workbook wb = Workbook.WB.Application.Workbooks[name + ".xlsm"];
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
