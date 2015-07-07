using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    /// <summary>
    /// Classe astratta che copre le funzionalità di aggiornamento della struttura e del riepilogo dei fogli di lavoro e del riepilogo.
    /// </summary>
    public abstract class AAggiorna
    {
        /// <summary>
        /// Launcher dell'aggiornamento in emergenza.
        /// </summary>
        public abstract void Emergenza();
        /// <summary>
        /// Aggiornamento dei fogli in emergenza.
        /// </summary>
        protected abstract void EmergenzaFogli();
        /// <summary>
        /// Aggiornamento del riepilogo in emergenza.
        /// </summary>
        protected abstract void EmergenzaRiepilogo();

        /// <summary>
        /// Launcher dell'aggiornamento dati.
        /// </summary>
        /// <returns></returns>
        public abstract bool Dati();
        /// <summary>
        /// Aggiornamento dei dati contenuti nei fogli.
        /// </summary>
        protected abstract void DatiFogli();
        /// <summary>
        /// Aggiornamento dei dati contenuti nel riepilogo.
        /// </summary>
        protected abstract void DatiRiepilogo();

        /// <summary>
        /// Launcher dell'aggiornamento della struttura.
        /// </summary>
        /// <returns></returns>
        public abstract bool Struttura();
        /// <summary>
        /// Aggiornamento della struttura dei fogli.
        /// </summary>
        protected abstract void StrutturaFogli();
        /// <summary>
        /// Aggiornamento della struttura del riepilogo.
        /// </summary>
        protected abstract void StrutturaRiepilogo();
    }

    /// <summary>
    /// Implementazione di base della classe AAggiorna. Nel caso nell'applicativo specifico ci fosse la necessità di variare la struttura di uno dei fogli, va fatto l'override dei questa classe ed eventualmente delle classi Scheet/Riepilogo a seconda del livello di personalizzazione.
    /// </summary>
    public class Aggiorna : AAggiorna
    {
        #region Costruttori

        public Aggiorna()
        {
            Workbook.Main.Select();
        }
        
        #endregion

        #region Metodi

        /// <summary>
        /// Launcher dell'aggiornamento della struttura.
        /// </summary>
        /// <returns>True se l'aggiornamento è andato a buon fine.</returns>
        public override bool Struttura()
        {
            if (DataBase.OpenConnection())
            {
                //aggiorno i parametri di base dell'applicazione
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
                        ws = Workbook.Sheets[categoria["DesCategoria"].ToString()];
                    }
                    catch
                    {
                        ws = (Excel.Worksheet)Workbook.Sheets.Add(Workbook.Log);
                        ws.Name = categoria["DesCategoria"].ToString();
                        ws.Select();
                        Workbook.Application.Windows[1].DisplayGridlines = false;
#if !DEBUG
                    Workbook.Application.ActiveWindow.DisplayHeadings = false;
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
                Workbook.Application.WindowState = Excel.XlWindowState.xlMaximized;
                
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
        /// <summary>
        /// Aggiornamento della struttura del riepilogo.
        /// </summary>
        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }
        /// <summary>
        /// Aggiornamento della struttura dei fogli.
        /// </summary>
        protected override void StrutturaFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.LoadStructure();
            }
        }

        /// <summary>
        /// Launcher dell'aggiornamento dati.
        /// </summary>
        /// <returns>True se l'aggiornamento è andato a buon fine.</returns>
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
        /// <summary>
        /// Aggiornamento dei dati contenuti nei fogli.
        /// </summary>
        protected override void DatiFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.UpdateData(true);
            }
        }
        /// <summary>
        /// Aggiornamento dei dati contenuti nel riepilogo.
        /// </summary>
        protected override void DatiRiepilogo()
        {
            Riepilogo main = new Riepilogo();
            main.UpdateData();
        }

        /// <summary>
        /// Launcher dell'aggiornamento in emergenza.
        /// </summary>
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
        /// <summary>
        /// Aggiornamento dei fogli in emergenza.
        /// </summary>
        protected override void EmergenzaFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.AggiornaDateTitoli();
            }
        }
        /// <summary>
        /// Aggiornamento del riepilogo in emergenza.
        /// </summary>
        protected override void EmergenzaRiepilogo()
        {
            Riepilogo main = new Riepilogo();
            main.RiepilogoInEmergenza();
        }

        #endregion
    }
}
