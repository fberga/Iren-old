using System;
using System.Collections.Generic;
using System.Drawing;
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

        protected virtual void ErrorStyle(ref TreeNode node)
        {
            node.BackColor = System.Drawing.Color.Red;
            node.ForeColor = System.Drawing.Color.Yellow;
            node.NodeFont = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold);
        }
        protected virtual void AlertStyle(ref TreeNode node)
        {
            node.BackColor = System.Drawing.Color.Yellow;
            node.ForeColor = System.Drawing.Color.Red;
            node.NodeFont = new System.Drawing.Font("Microsoft Sans Serif", 12, System.Drawing.FontStyle.Bold);
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
