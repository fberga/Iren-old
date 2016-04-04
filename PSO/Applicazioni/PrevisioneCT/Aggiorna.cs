using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;

namespace Iren.PSO.Applicazioni
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


        private void AggiornaCmbStagioni()
        {
            //seleziono la stagione nella combo
            Excel.Worksheet ws = Workbook.Sheets[DefinedNames.GetSheetName("CT_TORINO")];
            DefinedNames definedNames = new DefinedNames(ws.Name);
            Range rng = definedNames.Get("CT_TORINO", "STAGIONE", Date.SuffissoDATA1, Date.GetSuffissoOra(1));

            bool enabledEvents = Workbook.Application.EnableEvents;
            if(enabledEvents)
                Workbook.Application.EnableEvents = false;

            ((RibbonDropDown)Globals.Ribbons.GetRibbon<ToolsExcelRibbon>().Controls["cmbStagione"]).SelectedItemIndex = (int)(ws.Range[rng.ToString()].Value ?? 1) - 1;
            
            if(enabledEvents)
                Workbook.Application.EnableEvents = true;
        }

        public override bool Struttura(bool avoidRepositoryUpdate)
        {
            bool o = base.Struttura(avoidRepositoryUpdate);
            AggiornaCmbStagioni();
            return o;
        }

        public override bool Dati()
        {
            bool o = base.Dati();
            AggiornaCmbStagioni();
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
