using Iren.ToolsExcel.UserConfig;
using Iren.ToolsExcel.Core;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;


namespace Iren.ToolsExcel.Base
{
    public abstract class AEsporta
    {
        protected Core.DataBase _db = Utility.DataBase.DB;
        protected DataSet _localDB = Utility.DataBase.LocalDB;

        public abstract bool RunExport(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif);
        protected abstract bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif);

        protected Outlook.Application GetOutlookInstance()
        {
            Outlook.Application application = null;

            // Check whether there is an Outlook process running.
            if (Process.GetProcessesByName("OUTLOOK").Count() > 0)
            {

                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object.
                application = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            }
            else
            {

                // If not, create a new instance of Outlook and log on to the default profile.
                application = new Outlook.Application();
                Outlook.NameSpace nameSpace = application.GetNamespace("MAPI");
                nameSpace.Logon("", "");
                nameSpace = null;
            }

            // Return the Outlook Application object.
            return application;
        }

        protected virtual bool ExportToCSV(string nomeFile, DataTable dt)
        {
            if (dt.Rows.Count > 0)
            {
                try
                {
                    using (StreamWriter outFile = new StreamWriter(nomeFile))
                    {
                        foreach (DataRow r in dt.Rows)
                        {
                            IEnumerable<string> fields = r.ItemArray.Select(field => field.ToString());
                            outFile.WriteLine(string.Join(";", fields));
                        }
                        outFile.Flush();
                    }
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }

            return false;
        }
    }

    public class Esporta : AEsporta
    {

        public override bool RunExport(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            return true;
        }

        protected override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime dataRif)
        {
            return true;
        }
    }
}
