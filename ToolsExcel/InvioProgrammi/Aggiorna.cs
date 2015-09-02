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

        /// <summary>
        /// Aggiunta la creazione della struttura dei fogli di export.
        /// </summary>
        /// <returns>True se il processo va a buon fine.</returns>
        public override bool Struttura(bool avoidRepositoryUpdate)
        {
            //Aggiungo i fogli dei mercati leggendo da App.Config
            string[] mercati = Workbook.AppSettings("Mercati").Split('|');

            foreach (string msd in mercati)
            {
                Excel.Worksheet ws;
                try
                {
                    ws = Workbook.Sheets[msd];
                }
                catch
                {
                    ws = (Excel.Worksheet)Workbook.Sheets.Add(Workbook.Log);
                    ws.Name = msd;
                    ws.Select();
                    ws.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    Workbook.Application.Windows[1].DisplayGridlines = false;
                }
            }
            Workbook.Main.Select();
            Workbook.ScreenUpdating = false;

            return base.Struttura(avoidRepositoryUpdate);
        }
        /// <summary>
        /// Esegue prima la generazione dei fogli di export, successivamente quella dei fogli di lavoro.
        /// </summary>
        protected override void StrutturaFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.MSDSheets)
            {
                SheetExport se = new SheetExport(ws);
                se.LoadStructure();
            }

            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.LoadStructure();    
            }
        }
        /// <summary>
        /// I label sono diversi quindi viene utilizzato un init label customizzato.
        /// </summary>
        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }

        /// <summary>
        /// Aggiorna i dati dei fogli e dei fogli di export.
        /// </summary>
        /// <returns>True se il processo è andato a buon fine.</returns>
        public override bool Dati()
        {
            return base.Dati();
        }
        protected override void DatiFogli()
        {
            foreach (Excel.Worksheet ws in Workbook.MSDSheets)
            {
                SheetExport se = new SheetExport(ws);
                se.UpdateData();
            }

            foreach (Excel.Worksheet ws in Workbook.CategorySheets)
            {
                Sheet s = new Sheet(ws);
                s.UpdateData();
            }
        }
    }
}
