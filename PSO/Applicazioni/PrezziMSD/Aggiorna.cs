using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Applicazioni
{
    /// <summary>
    /// Label riepilogo custom.
    /// </summary>
    public class Aggiorna : Base.Aggiorna
    {
        public Aggiorna()
            : base()
        {

        }        

        protected override void StrutturaRiepilogo()
        {
            Riepilogo riepilogo = new Riepilogo();
            riepilogo.LoadStructure();
        }
    }

}
