using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.FrontOffice.Base
{
    internal class Struttura
    {
        public int colBlock = 5,
            rigaBlock = 6,
            rigaGoto = 3,
            intervalloGiorni = 0,
            colRecap = 165,
            rowRecap = 2;
        public bool visData0H24 = false,
            visParametro = false;

        public Struttura() { }
    }
}
