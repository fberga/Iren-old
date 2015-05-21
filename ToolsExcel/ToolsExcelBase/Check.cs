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
        #region Variabili

        protected Excel.Worksheet _ws;
        protected DefinedNames _nomiDefiniti;
        protected CheckObj _check;

        #endregion

        #region Metodi

        public virtual CheckOutput ExecuteCheck(Excel.Worksheet ws, DefinedNames definedNames, CheckObj check)
        {
            return new CheckOutput();
        }

        protected decimal GetDecimal(object siglaEntita, object siglaInformazione, object suffissoData, object suffissoOra)
        {
            return (decimal)(_ws.Range[_nomiDefiniti.Get(siglaEntita, siglaInformazione, suffissoData, suffissoOra).ToString()].Value ?? 0);
        }
        protected object GetObject(object siglaEntita, object siglaInformazione, object suffissoData, object suffissoOra)
        {
            return _ws.Range[_nomiDefiniti.Get(siglaEntita, siglaInformazione, suffissoData, suffissoOra).ToString()].Value;
        }
        protected string GetString(object siglaEntita, object siglaInformazione, object suffissoData, object suffissoOra)
        {
            return (string)(_ws.Range[_nomiDefiniti.Get(siglaEntita, siglaInformazione, suffissoData, suffissoOra).ToString()].Value ?? "");
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

        #endregion
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

    public class CheckObj
    {
        #region Proprietà

        public string SiglaEntita { get; set; }
        public string Range { get; set; }
        public int Type { get; set; }

        #endregion

        #region Costruttori

        public CheckObj(string range)
        {
            Range = range;
        }
        public CheckObj(string siglaEntita, string range, int type)
        {
            SiglaEntita = siglaEntita;
            Range = range;
            Type = type;
        }

        #endregion

        #region Metodi

        #endregion
    }
}

