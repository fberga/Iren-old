using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Base
{
    public class Check
    {
        public virtual TreeNode ExecuteCheck(Excel.Worksheet ws, NewDefinedNames newNomiDefiniti, CheckObj check)
        {
            return new TreeNode();
        }  
    }
}
