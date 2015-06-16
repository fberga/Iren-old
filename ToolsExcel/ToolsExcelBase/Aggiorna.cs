using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    public abstract class AAggiorna
    {
        public abstract void Emergenza();
        protected abstract void EmergenzaFogli();
        protected abstract void EmergenzaRiepilogo();

        public abstract bool Dati();
        protected abstract void DatiFogli();
        protected abstract void DatiRiepilogo();

        public abstract bool Struttura();
        protected abstract void StrutturaFogli();
        protected abstract void StrutturaRiepilogo();
    }

    public class Aggiorna : AAggiorna
    {
        public Aggiorna()
        {
            Workbook.Main.Select();
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
        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }
        protected override void StrutturaFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.Sheets)
            {
                Sheet s = new Sheet(ws);
                s.LoadStructure();
            }
        }

        public override bool Dati()
        {
            if (DataBase.OpenConnection())
            {
                SplashScreen.Show();

                bool wasProtected = Sheet.Protected;
                if (wasProtected)
                    Sheet.Protected = false;

                Workbook.ScreenUpdating = false;

                SplashScreen.UpdateStatus("Aggiorno dati Riepilogo");
                DatiRiepilogo();
                SplashScreen.UpdateStatus("Aggiorno dati Fogli");
                DatiFogli();

                if (wasProtected)
                    Sheet.Protected = true;
                Workbook.ScreenUpdating = true;
                SplashScreen.Close();

                return true;
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("Impossibile aggiornare i dati: ci sono problemi di connessione o la funzione Forza Emergenza è attiva.", Simboli.nomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                return false;
            }
        }
        protected override void DatiFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.Sheets)
            {
                Sheet s = new Sheet(ws);
                s.UpdateData(true);
            }
        }
        protected override void DatiRiepilogo()
        {
            Riepilogo main = new Riepilogo();
            main.UpdateData();
        }

        public override void Emergenza()
        {
            SplashScreen.Show();

            bool wasProtected = Sheet.Protected;
            if (wasProtected)
                Sheet.Protected = false;

            Workbook.ScreenUpdating = false;

            SplashScreen.UpdateStatus("Riepilogo in emergenza");
            EmergenzaRiepilogo();
            SplashScreen.UpdateStatus("Aggiorno le date");
            EmergenzaFogli();

            if (wasProtected)
                Sheet.Protected = true;
            Workbook.ScreenUpdating = true;
            SplashScreen.Close();
        }

        protected override void EmergenzaFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.Sheets)
            {
                Sheet s = new Sheet(ws);
                s.AggiornaDateTitoli();
            }
        }
        protected override void EmergenzaRiepilogo()
        {
            Riepilogo main = new Riepilogo();
            main.RiepilogoInEmergenza();
        }

    }
}
