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
        Microsoft.Office.Tools.Excel.Worksheet ActiveSheet { get; }
        Microsoft.Office.Tools.Excel.Worksheet Main { get; }
        Microsoft.Office.Tools.Excel.Worksheet Log { get; }
        Excel.Sheets Sheets { get; }
        Excel.Application Application { get; }

        string Name { get; }
        string Path { get; }
        string FullName { get; }

        int IdApplicazione { get; set; }
        int IdUtente { get; set; }
        string NomeUtente { get; set; }
        DateTime DataAttiva { get; set; }
        string Ambiente { get; set; }

        DataSet RepositoryDataSet { get; }
        DataSet LogDataSet { get; }
        DataSet RibbonDataSet { get; }
    }
}
