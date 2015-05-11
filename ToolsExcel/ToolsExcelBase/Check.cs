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
        public virtual CheckOutput ExecuteCheck(Excel.Worksheet ws, NewDefinedNames newNomiDefiniti, CheckObj check)
        {
            return new CheckOutput();
        }  
    }

    public class CheckOutput
    {
        public enum CheckStatus
        {
            Ok, Alert, Error
        }

        TreeNode _node;
        CheckStatus _status;

        public CheckOutput()
        {
            _node = new TreeNode();
            _status = CheckStatus.Ok;
        }

        public CheckOutput(TreeNode node, CheckStatus status)
        {
            _node = node;
            _status = status;
        }

        public TreeNode Node { get { return _node; } }
        public CheckStatus Status { get { return _status; } }
    }
}
