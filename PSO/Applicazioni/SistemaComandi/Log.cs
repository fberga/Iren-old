using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Core;
using Iren.PSO;
using Iren.PSO.Base;

namespace Iren.PSO.Applicazioni
{
    public partial class Log
    {
        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(this.Log_Startup);
        }

        #endregion

        #region Callbacks

        private void Log_Startup(object sender, EventArgs e)
        {
            Unprotect(PSO.Base.Workbook.Password);

            this.logList.DataSource = PSO.Base.Workbook.LogDataTable;

            this.logList.AutoSetDataBoundColumnHeaders = true;
            this.logList.Range.EntireColumn.AutoFit();
            this.logList.TableStyle = "TableStyleLight16";

            ((Excel.Range)Columns[2]).NumberFormat = "dd/MM/yyyy";
            ((Excel.Range)Columns[2]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            Protect(PSO.Base.Workbook.Password, allowSorting: true, allowFiltering: true);
        }

        #endregion

    }
}
