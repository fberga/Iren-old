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
            //controllo che la selezione non sia multi-linea con in mezzo delle righe nascoste - nel caso avverto l'utente che non può effettuare modifiche
            if (Target.Rows.Count > 1)
            {
                if(Simboli.ModificaDati)
                {
                    foreach (Excel.Range r in Target.Rows)
                    {
                        if (r.EntireRow.Hidden)
                        {
                            System.Windows.Forms.MessageBox.Show("Nella selezione sono incluse righe nascoste. Non si può procedere con la modifica...", Simboli.nomeApplicazione + " - ATTENZIONE", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Stop);

                            Target.Cells[1, 1].Select();

                            break;
                        }
                    }
                }    
            }
            else
            {
                try
                {
                    NewDefinedNames newDefinedNames = new NewDefinedNames(Target.Worksheet.Name, NewDefinedNames.InitType.GOTOsOnly);
                    string address = newDefinedNames.GetGOTO(Range.R1C1toA1(Target.Row, Target.Column));
                    Goto(address);
                }
                catch { }
            }
        }

        /// <summary>
        /// Sposta la selezione su address e la centra nello schermo.
        /// </summary>
        /// <param name="address">L'indirizzo della cella/range da selezionare in forma A1</param>
        public static void Goto(string address)
        {
            if (address != "")
            {
                Excel.Range rng = (Excel.Range)Workbook.WB.Application.Range[address];
                rng.Worksheet.Activate();
                rng.Select();
                Workbook.WB.Application.ActiveWindow.SmallScroll(rng.Row - Workbook.WB.Application.ActiveWindow.VisibleRange.Cells[1, 1].Row - 1);
            }
        }

        public static void StoreEdit(object Sh, Excel.Range Target)
        {
            bool wasProtected = Target.Worksheet.ProtectContents;
            bool screenUpdating = Target.Application.ScreenUpdating;
            if (wasProtected)
                Target.Worksheet.Unprotect(Simboli.pwd);
            
            if (screenUpdating)
                Target.Application.ScreenUpdating = false;

            NewDefinedNames newNomiDefiniti = new NewDefinedNames(Target.Worksheet.Name, NewDefinedNames.InitType.SaveDB);
            DataTable modifiche = DataBase.LocalDB.Tables[DataBase.Tab.MODIFICA];
            

            Excel.Worksheet ws = (Excel.Worksheet)Sh;
            if (ws.ChartObjects().Count > 0)
            {
                Sheet s = new Sheet(ws);
                s.AggiornaGrafici();
            }

            string[] ranges = Target.Address.Split(',');
            
            foreach (string range in ranges)
            {
                Range rng = new Range(range);
                foreach (Range row in rng.Rows)
                {
                    if (newNomiDefiniti.SaveDB(row.StartRow))
                    {
                        bool annota = newNomiDefiniti.ToNote(row.StartRow) && !Workbook.DaElaborazione;
                        foreach (Range column in row.Columns)
                        {
                            string[] parts = newNomiDefiniti.GetNameByAddress(column.StartRow, column.StartColumn).Split(Simboli.UNION[0]);
                            string data;
                            if(parts.Length == 4)
                                data = Date.GetDataFromSuffisso(parts[2], parts[3]);
                            else
                                data = Date.GetDataFromSuffisso(parts[2], "");

                            DataRow r = modifiche.Rows.Find(new object[] { parts[0], parts[1], data});
                            if (r != null)
                                r["Valore"] = ws.Range[column.ToString()].Value;
                            else
                            {
                                DataRow newRow = modifiche.NewRow();

                                newRow["SiglaEntita"] = parts[0];
                                newRow["SiglaInformazione"] = parts[1];
                                newRow["Data"] = data;
                                newRow["Valore"] = ws.Range[column.ToString()].Value;
                                newRow["AnnotaModifica"] = annota ? "1" : "0";
                                newRow["IdApplicazione"] = DataBase.DB.IdApplicazione;
                                newRow["IdUtente"] = DataBase.DB.IdUtenteAttivo;

                                modifiche.Rows.Add(newRow);
                            }

                            if (annota)
                            {
                                ws.Range[column.ToString()].ClearComments();
                                ws.Range[column.ToString()].AddComment("Valore inserito manualmente").Visible = false;
                            }
                        }
                    }
                }
            }

            if (wasProtected)
                Target.Worksheet.Protect(Simboli.pwd);
            
           if (screenUpdating)
                Target.Application.ScreenUpdating = true;
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
