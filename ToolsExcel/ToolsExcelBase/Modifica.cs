﻿using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    /// <summary>
    /// Interfaccia per implementare delle funzioni particolari quando l'utente entra in modalità di modifica.
    /// </summary>
    public abstract class AModifica
    {
        /// <summary>
        /// Handler per l'evento SheetChange che viene aggiunto quando l'utente va in modifica e permette di definire azioni custom (i.e. copiare il dato nel foglio MSD corretto).
        /// </summary>
        /// <param name="Sh">Sheet</param>
        /// <param name="Target">Microsoft.Office.Interop.Excel.Range dove avviene la modifica.</param>
        public abstract void Range(object Sh, Excel.Range Target);
    }

    public class Modifica : AModifica
    {
        /// <summary>
        /// Handler per l'evento SheetChange che viene aggiunto quando l'utente va in modifica e permette di definire azioni custom (i.e. copiare il dato nel foglio MSD corretto).
        /// </summary>
        /// <param name="Sh">Sheet</param>
        /// <param name="Target">Microsoft.Office.Interop.Excel.Range dove avviene la modifica.</param>
        public override void Range(object Sh, Excel.Range Target)
        {
            return;
        }
    }
}
