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
            if (base.Struttura())
            {
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

                //


                return true;
            }

            return false;
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
        }

        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }
    }

}
