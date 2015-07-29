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
    /// <summary>
    /// Classe che gestisce molti dei comportamenti standard del workbook.
    /// </summary>
    public class Handler
    {
        #region Metodi 

        /// <summary>
        /// Handler per il click su celle di selezione.
        /// </summary>
        /// <param name="Sh"></param>
        /// <param name="Target"></param>
        public static void SelectionClick(object Sh, Excel.Range Target)
        {
            Workbook.ScreenUpdating = false;
            DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.Selection);
            Range rng = new Range(Target.Row, Target.Column);
            Selection sel;
            int val;
            if (definedNames.HasSelections() && definedNames.TryGetSelectionByPeer(rng, out sel, out val))
            {
                Target.Worksheet.Unprotect(Simboli.pwd);
                if (sel != null)
                {
                    sel.ClearSelections(Target.Worksheet);
                    sel.Select(Target.Worksheet, rng.ToString());

                    //annoto modifiche e le salvo sul DB
                    Target.Worksheet.Range[sel.RifAddress].Value = val;
                    DataBase.SalvaModificheDB();
                }
                Target.Worksheet.Protect(Simboli.pwd);
            }
            Workbook.ScreenUpdating = true;
        }
        /// <summary>
        /// Gestisce il caso in cui ci sia una selezione multipla che andrebbe a scrivere su righe nascoste: allerta l'utente e impedisce di procedere con la modifica.
        /// </summary>
        /// <param name="Sh">Sheet di provenienza.</param>
        /// <param name="Target">Range selezionato dall'utente.</param>
        public static void CellClick(object Sh, Excel.Range Target)
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
                    DefinedNames newDefinedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.GOTOs);
                    string address = newDefinedNames.GetGotoFromAddress(Range.R1C1toA1(Target.Row, Target.Column));
                    Goto(address);
                }
                catch {}
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
                Excel.Range rng = (Excel.Range)Workbook.Application.Range[address];
                rng.Worksheet.Activate();
                rng.Select();
                Workbook.Application.ActiveWindow.SmallScroll(rng.Row - Workbook.Application.ActiveWindow.VisibleRange.Cells[1, 1].Row - 1);
            }
        }
        /// <summary>
        /// Funzione per il salvataggio delle modifiche apportate a ranges anche non contigui.
        /// </summary>
        /// <param name="Target">L'insieme dei ranges modificati</param>
        /// <param name="annotaModifica">Se la modifica va segnalata all'utente attraverso il commento sulla cella oppure no.</param>
        /// <param name="fromCalcolo">Flag per eseguire azioni particolari nel caso la provenienza del salvataggio sia da un calcolo.</param>
        public static void StoreEdit(Excel.Range Target, int annotaModifica = -1, bool fromCalcolo = false)
        {
            Excel.Worksheet ws = Target.Worksheet;
            bool wasProtected = ws.ProtectContents;
            bool screenUpdating = ws.Application.ScreenUpdating;
            if (wasProtected)
                ws.Unprotect(Simboli.pwd);

            if (screenUpdating)
               Workbook.ScreenUpdating = false;

            DefinedNames definedNames = new DefinedNames(Target.Worksheet.Name, DefinedNames.InitType.SaveDB);
            DataTable modifiche = DataBase.LocalDB.Tables[DataBase.Tab.MODIFICA];

            if (ws.ChartObjects().Count > 0 && !fromCalcolo)
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
                    if (definedNames.SaveDB(row.StartRow))
                    {
                        bool annota = annotaModifica == -1 ? definedNames.ToNote(row.StartRow) : annotaModifica == 1;
                        foreach (Range column in row.Columns)
                        {
                            string[] parts = definedNames.GetNameByAddress(column.StartRow, column.StartColumn).Split(Simboli.UNION[0]);

                            string data;
                            if (parts.Length == 4)
                                data = Date.GetDataFromSuffisso(parts[2], parts[3]);
                            else
                                data = Date.GetDataFromSuffisso(parts[2], "");

                            DataRow r = modifiche.Rows.Find(new object[] { parts[0], parts[1], data });
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
                ws.Protect(Simboli.pwd);

            if (screenUpdating)
                ws.Application.ScreenUpdating = true;
        }
        /// <summary>
        /// Funzione per il salvataggio delle modifiche apportate dall'utente quando la modifica è abilitata.
        /// </summary>
        /// <param name="Sh">Sheet.</param>
        /// <param name="Target">Range.</param>
        public static void StoreEdit(object Sh, Excel.Range Target)
        {
            StoreEdit(Target);
        }
        /// <summary>
        /// Handler per cambiare il label di modifica e la scritta sotto il tasto sul ribbon.
        /// </summary>
        /// <param name="modifica">True se modifica è abilitata, false se disabilitata.</param>
        public static void ChangeModificaDati(bool modifica)
        {
            Excel.Worksheet ws = Workbook.Sheets["Main"];
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
        /// <summary>
        /// Handler per cambiare il label dell'ambiente.
        /// </summary>
        /// <param name="ambiente">Sigla Ambiente.</param>
        public static void ChangeAmbiente(string ambiente)
        {
            Workbook.Main.Shapes.Item("lbTest").Locked = false;
            switch (ambiente)
            {
                case "Dev":
                    Workbook.Main.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: DEVELOPMENT";
                    //rosso
                    Workbook.Main.Shapes.Item("lbTest").Line.Weight = 2f;
                    Workbook.Main.Shapes.Item("lbTest").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(140, 56, 54));
                    Workbook.Main.Shapes.Item("lbTest").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 80, 77));
                    break;
                case "Test":
                    Workbook.Main.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: TEST";
                    //giallo
                    Workbook.Main.Shapes.Item("lbTest").Line.Weight = 2f;
                    Workbook.Main.Shapes.Item("lbTest").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 204, 0));
                    Workbook.Main.Shapes.Item("lbTest").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 102));
                    break;
                case "Produzione":
                    Workbook.Main.Shapes.Item("lbTest").TextFrame.Characters().Text = "Ambiente: PRODUZIONE";
                    //bianco normale
                    Workbook.Main.Shapes.Item("lbTest").Line.Weight = 0.75f;
                    Workbook.Main.Shapes.Item("lbTest").Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 255, 255));
                    Workbook.Main.Shapes.Item("lbTest").Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(0, 0, 0));
                    Workbook.Main.Shapes.Item("lbTest").Line.ForeColor.Brightness = +0.75f;
                    break;
            }
            Workbook.Main.Shapes.Item("lbTest").Locked = true;
        }
        /// <summary>
        /// Handler per cambiare i label in base alla modifica dello stato del DB.
        /// </summary>
        /// <param name="db">Database interessato</param>
        /// <param name="online">True se il database è online, false altrimenti.</param>
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

            Excel.Worksheet ws = Workbook.Main;
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
        /// <summary>
        /// Handler per la modifica del label che indica il mercato attivo.
        /// </summary>
        /// <param name="mercato">La stringa con il nome del mercato.</param>
        public static void ChangeMercatoAttivo(string mercato)
        {
            Workbook.Main.Shapes.Item("lbMercato").Locked = false;
            Workbook.Main.Shapes.Item("lbMercato").TextFrame.Characters().Text = mercato;
            Workbook.Main.Shapes.Item("lbMercato").Locked = true;
        }
        /// <summary>
        /// Gestisce l'apertura degli altri file con i tasti sul ribbon.
        /// </summary>
        /// <param name="name">Nome dell'applicazione da aprire.</param>
        public static void SwitchWorksheet(string name)
        {
            //TODO aprire gli altri file!!!!!!
            string path = "";
#if(DEBUG)
            path = "D:\\Repository\\Iren\\ToolsExcel\\"+name+"\\bin\\Debug\\"+name+".xlsm";
#else
            path = ".\\"+name+".xlsm";
#endif
            if (!IsWorkbookOpen(name))
            {
                Workbook.Application.Workbooks.Open(path);
            }
            else
            {
                try
                {
                    Workbook.Application.Workbooks[name + ".xlsm"].Activate();
                }
                catch(Exception)
                {
                    Workbook.Application.Workbooks[name].Activate();
                }
            }
        }
        /// <summary>
        /// Verifica se il workbook indicato da nome è aperto oppure no.
        /// </summary>
        /// <param name="name">Nome del workbook.</param>
        /// <returns>True se è aperto, false altrimenti.</returns>
        private static bool IsWorkbookOpen(string name)
        {
            try
            {
                Microsoft.Office.Interop.Excel._Workbook wb = Workbook.Application.Workbooks[name + ".xlsm"];
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion
    }
}
