using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    /// <summary>
    /// Aggiungo la stagione al foglio e carico la struttura del riepilogo personalizzata.
    /// </summary>
    public class Aggiorna : Base.Aggiorna
    {
        public Aggiorna()
            : base()
        {

        }

        public override bool Struttura(bool avoidRepositoryUpdate)
        {
            bool o = base.Struttura(avoidRepositoryUpdate);

            //forzo aggiornamento della stagione
            Handler.ScriviStagione(Workbook.IdStagione);

            return o;
        }

        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }

        protected override void DatiRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.UpdateData();
        }
    }

}
