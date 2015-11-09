using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Base
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
        public abstract bool Struttura(bool avoidRepositoryUpdate);
        /// <summary>
        /// Aggiornamento della struttura dei fogli.
        /// </summary>
        protected abstract void StrutturaFogli();
        /// <summary>
        /// Aggiornamento della struttura del riepilogo.
        /// </summary>
        protected abstract void StrutturaRiepilogo();

        public abstract void ColoriData();
    }

    /// <summary>
    /// Implementazione di base della classe AAggiorna. Nel caso nell'applicativo specifico ci fosse la necessità di variare la struttura di uno dei fogli, va fatto l'override dei questa classe ed eventualmente delle classi Scheet/Riepilogo a seconda del livello di personalizzazione.
    /// </summary>
    public class Aggiorna : AAggiorna
    {
        #region Variabili

        public static Dictionary<string, Tuple<int, int>> _freezePanes = new Dictionary<string, Tuple<int, int>>();

        #endregion


        #region Costruttori

        public Aggiorna()
        {
            //Workbook.Main.Select();
        }
        
        #endregion

        #region Metodi

        /// <summary>
        /// Carica tutti i valori dal DB.
        /// </summary>
        protected void CaricaDatiDalDB()
        {
            CancellaTabelle();
            DataTable entitaProprieta = Workbook.Repository[DataBase.TAB.ENTITA_PROPRIETA];
            DateTime dataFine = Workbook.DataAttiva.AddDays(Math.Max(
                    (from r in entitaProprieta.AsEnumerable()
                     where r["IdApplicazione"].Equals(Workbook.IdApplicazione) && r["SiglaProprieta"].ToString().EndsWith("GIORNI_STRUTTURA")
                     select int.Parse(r["Valore"].ToString())).DefaultIfEmpty().Max(), Struct.intervalloGiorni));

            SplashScreen.UpdateStatus("Carico informazioni dal DB");
            DataTable datiApplicazioneH = DataBase.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_H, "@SiglaCategoria=ALL;@SiglaEntita=ALL;@DateFrom=" + Workbook.DataAttiva.ToString("yyyyMMdd") + ";@DateTo=" + dataFine.ToString("yyyyMMdd") + ";@Tipo=1;@All=1") ?? new DataTable();

            datiApplicazioneH.TableName = DataBase.TAB.DATI_APPLICAZIONE_H;
            Workbook.Repository.Add(datiApplicazioneH);

            SplashScreen.UpdateStatus("Carico commenti dal DB");
            DataTable insertManuali = DataBase.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_COMMENTO, "@SiglaCategoria=ALL;@SiglaEntita=ALL;@DateFrom=" + Workbook.DataAttiva.ToString("yyyyMMdd") + ";@DateTo=" + dataFine.ToString("yyyyMMdd") + ";@All=1") ?? new DataTable();

            insertManuali.TableName = DataBase.TAB.DATI_APPLICAZIONE_COMMENTO;
            Workbook.Repository.Add(insertManuali);

            SplashScreen.UpdateStatus("Carico informazioni giornaliere dal DB");
            DataTable datiApplicazioneD = DataBase.Select(DataBase.SP.APPLICAZIONE_INFORMAZIONE_D, "@SiglaCategoria=ALL;@SiglaEntita=ALL;@DateFrom=" + Workbook.DataAttiva.ToString("yyyyMMdd") + ";@DateTo=" + Workbook.DataAttiva.AddDays(Struct.intervalloGiorni).ToString("yyyyMMdd") + ";@Tipo=1;@All=1") ?? new DataTable();

            datiApplicazioneD.TableName = DataBase.TAB.DATI_APPLICAZIONE_D;
            Workbook.Repository.Add(datiApplicazioneD);
        }
        /// <summary>
        /// Cancella le tabelle create in modo da non avere duplicati nel dataset.
        /// </summary>
        protected void CancellaTabelle()
        {
            //elimino le tabelle con le informazioni ormai scritte nel foglio
            if (Workbook.Repository.Contains(DataBase.TAB.DATI_APPLICAZIONE_H))
                Workbook.Repository.Remove(DataBase.TAB.DATI_APPLICAZIONE_H);
            if (Workbook.Repository.Contains(DataBase.TAB.DATI_APPLICAZIONE_D))
                Workbook.Repository.Remove(DataBase.TAB.DATI_APPLICAZIONE_D);
            if (Workbook.Repository.Contains(DataBase.TAB.DATI_APPLICAZIONE_COMMENTO))
                Workbook.Repository.Remove(DataBase.TAB.DATI_APPLICAZIONE_COMMENTO);
        }

        /// <summary>
        /// Launcher dell'aggiornamento della struttura.
        /// </summary>
        /// <returns>True se l'aggiornamento è andato a buon fine.</returns>
        public override bool Struttura(bool avoidRepositoryUpdate)
        {
            if (DataBase.OpenConnection() || avoidRepositoryUpdate)
            {
                //aggiorno i parametri di base dell'applicazione
                if (!avoidRepositoryUpdate)
                    Workbook.AggiornaParametriApplicazione();

                SplashScreen.Show();

                bool wasProtected = Sheet.Protected;
                if (wasProtected)
                    Sheet.Protected = false;

                Workbook.ScreenUpdating = false;

                if (!avoidRepositoryUpdate)
                {
                    SplashScreen.UpdateStatus("Carico struttura dal DB");
                    Workbook.Repository.Aggiorna();
                }
                else
                {
                    //resetto la struttura nomi che verrà ricreata nelle prossime righe
                    Workbook.Repository.InitStrutturaNomi();
                }
                
                SplashScreen.UpdateStatus("Controllo se tutti i fogli sono presenti");

                DataView categorie = Workbook.Repository[DataBase.TAB.CATEGORIA].DefaultView;
                categorie.RowFilter = "Operativa = 1 AND IdApplicazione = " + Workbook.IdApplicazione;

                foreach (DataRowView categoria in categorie)
                {
                    Excel.Worksheet ws;
                    try
                    {
                        ws = Workbook.Sheets[categoria["DesCategoria"].ToString()];
                        ws.Activate();
                        if(_freezePanes.ContainsKey(ws.Name))
                            _freezePanes[ws.Name] = Tuple.Create<int,int>(Workbook.Application.ActiveWindow.SplitRow + 1, Workbook.Application.ActiveWindow.SplitColumn + 1);
                        else
                            _freezePanes.Add(ws.Name, Tuple.Create<int,int>(Workbook.Application.ActiveWindow.SplitRow + 1, Workbook.Application.ActiveWindow.SplitColumn + 1));
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
                Workbook.ScreenUpdating = false;

                try
                {
                    if(DataBase.OpenConnection())
                        CaricaDatiDalDB();

                    Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
                    SplashScreen.UpdateStatus("Aggiorno struttura Riepilogo");
                    StrutturaRiepilogo();

                    SplashScreen.UpdateStatus("Aggiorno struttura Fogli");
                    StrutturaFogli();

                    SplashScreen.UpdateStatus("Salvo struttura in locale");
                    //Workbook.DumpDataSet();

                    SplashScreen.UpdateStatus("Invio modifiche al server");
                    Workbook.ScreenUpdating = false;
                    Sheet.SalvaModifiche();
                    DataBase.SalvaModificheDB();

                    SplashScreen.UpdateStatus("Azzero selezioni");
                    foreach (Excel.Worksheet ws in Workbook.Sheets)
                    {
                        if (ws.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            ws.Activate();
                            ws.Range["A1"].Select();
                        }
                    }

                    Workbook.Main.Select();
                    SplashScreen.UpdateStatus("Abilito calcolo automatico");
                    Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    Workbook.Application.WindowState = Excel.XlWindowState.xlMaximized;

                    if (wasProtected)
                        Sheet.Protected = true;

                    SplashScreen.Close();
                    CancellaTabelle();
                    return true;
                }
                catch
                {
                    SplashScreen.Close();
                    CancellaTabelle();
                    return false;
                }
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

                try
                {
                    CaricaDatiDalDB();
                    Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
                    SplashScreen.UpdateStatus("Aggiorno dati Riepilogo");
                    DatiRiepilogo();
                    SplashScreen.UpdateStatus("Aggiorno dati Fogli");
                    DatiFogli();

                    SplashScreen.UpdateStatus("Invio modifiche al server");
                    Workbook.ScreenUpdating = false;
                    Sheet.SalvaModifiche();
                    DataBase.SalvaModificheDB();

                    SplashScreen.UpdateStatus("Abilito calcolo automatico");
                    Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                    if (wasProtected)
                        Sheet.Protected = true;
                    Workbook.ScreenUpdating = true;
                    SplashScreen.Close();

                    CancellaTabelle();
                }
                catch
                {
                    Workbook.ScreenUpdating = true;
                    SplashScreen.Close();

                    CancellaTabelle();
                    return false;
                }

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
                s.UpdateData();
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

        public override void ColoriData()
        {
            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.UpdateDayColor();
            }
        }

        #endregion
    }
}
