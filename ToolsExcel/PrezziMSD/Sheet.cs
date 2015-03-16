using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel
{
    class Sheet : Base.Sheet
    {
        protected override void InsertInformazioniEntita(object siglaEntita)
        {
            DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITAPROPRIETA].DefaultView;
            entitaProprieta.RowFilter = "SiglaEntita = '" + siglaEntita + "' AND SiglaProprieta = 'MSD_ACCENSIONE'";

            if (entitaProprieta.Count > 0)
            {
                //dalla classe base mi trovo in _rigaAttiva = prima riga di informazioni dell'entità
            }
        }
    }
}
