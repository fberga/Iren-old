using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;

namespace Iren.FrontOffice.Automation
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = false;
            xlApp.Workbooks.Open(ConfigurationManager.AppSettings["path"]);
            xlApp.Run(ConfigurationManager.AppSettings["macro"]);
            xlApp.Quit();
        }
    }
}
