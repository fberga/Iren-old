using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    public class Aggiorna : Base.Aggiorna
    {
        public Aggiorna()
            : base()
        {

        }

        public override bool Struttura()
        {
            if (DataBase.OpenConnection())
            {
                Workbook.AggiornaParametriApplicazione();

                SplashScreen.Show();

                bool wasProtected = Sheet.Protected;
                if (wasProtected)
                    Sheet.Protected = false;

                Workbook.ScreenUpdating = false;

                SplashScreen.UpdateStatus("Carico struttura dal DB");
                Repository.Aggiorna();

                SplashScreen.UpdateStatus("Controllo se tutti i fogli sono presenti");

                DataView categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
                categorie.RowFilter = "Operativa = 1";

                foreach (DataRowView categoria in categorie)
                {
                    Excel.Worksheet ws;
                    try
                    {
                        ws = Workbook.WB.Worksheets[categoria["DesCategoria"].ToString()];
                    }
                    catch
                    {
                        ws = (Excel.Worksheet)Workbook.WB.Worksheets.Add(Workbook.Log);
                        ws.Name = categoria["DesCategoria"].ToString();
                        ws.Select();
                        Workbook.WB.Application.Windows[1].DisplayGridlines = false;
#if !DEBUG
                    Workbook.WB.Application.ActiveWindow.DisplayHeadings = false;
#endif
                    }
                }

                //Aggiungo i fogli dei mercati leggendo da App.Config
                string[] mercati = Workbook.AppSettings("Mercati").Split('|');

                foreach (string msd in mercati)
                {
                    Excel.Worksheet ws;
                    try
                    {
                        ws = Workbook.WB.Worksheets[msd];
                    }
                    catch
                    {
                        ws = (Excel.Worksheet)Workbook.WB.Worksheets.Add(Workbook.Log);
                        ws.Name = msd;
                        ws.Select();
                        Workbook.WB.Application.Windows[1].DisplayGridlines = false;
                    }
                }

                SplashScreen.UpdateStatus("Aggiorno struttura Riepilogo");
                StrutturaRiepilogo();

                SplashScreen.UpdateStatus("Aggiorno struttura Fogli");
                StrutturaFogli();

                SplashScreen.UpdateStatus("Salvo struttura in locale");
                Workbook.DumpDataSet();

                Workbook.Main.Select();
                Workbook.Main.Range["A1"].Select();
                Workbook.WB.Application.WindowState = Excel.XlWindowState.xlMaximized;

                if (wasProtected)
                    Sheet.Protected = true;

                Workbook.ScreenUpdating = true;
                SplashScreen.Close();

                return true;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Impossibile aggiornare la struttura: ci sono problemi di connessione o la funzione Forza Emergenza è attiva.", Simboli.nomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                return false;
            }
        }

        protected override void StrutturaFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.Sheets)
            {
                if (!ws.Name.StartsWith("MSD"))
                {
                    Sheet s = new Sheet(ws);
                    s.LoadStructure();
                }
            }

            foreach (Excel.Worksheet ws in Workbook.Sheets)
            {
                if (ws.Name.StartsWith("MSD"))
                {
                    SheetExport s = new SheetExport(ws);
                    s.LoadStructure();
                }
            }
        }

        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }
    }

}
