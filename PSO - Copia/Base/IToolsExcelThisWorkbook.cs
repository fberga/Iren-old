using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.ToolsExcel.Base
{
    public interface IToolsExcelThisWorkbook
    {
        System.Version Version { get; }
        Microsoft.Office.Tools.Excel.Workbook Base { get; }
        Microsoft.Office.Interop.Excel.Worksheet ActiveSheet { get; }        
        Excel.Sheets Sheets { get; }
        Excel.Application Application { get; }

        string Name { get; }
        string Path { get; }
        string FullName { get; }
        string Pwd { get; }
        string NomeUtente { get; set; }
        string Ambiente { get; set; }

        int IdApplicazione { get; set; }
        int IdUtente { get; set; }
        int IdStagione { get; set; }

        DateTime DataAttiva { get; set; }

        System.Data.DataSet RepositoryDataSet { get; }
        System.Data.DataTable LogDataTable { get; set; }
        System.Data.DataSet RibbonDataSet { get; }
    }
}
