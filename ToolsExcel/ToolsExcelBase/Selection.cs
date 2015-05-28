using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.Base
{
    public class Selection
    {
        #region Variabili
        
        private string _rif = "";
        private Dictionary<string, int> _peers = new Dictionary<string, int>();
        
        #endregion

        #region Proprietà

        public string RifAddress { get { return _rif; } }
        public Dictionary<string, int> SelPeers { get { return _peers; } }

        #endregion

        #region Costruttore

        public Selection(string rifAddress, Dictionary<string, int> peers)
        {
            _rif = rifAddress;
            _peers = peers;
        }

        #endregion

        #region Metodi

        public void ClearSelections(Microsoft.Office.Interop.Excel.Worksheet ws)
        {
            foreach (string cell in SelPeers.Keys)
            {
                double height = ws.Range[cell].RowHeight;
                ws.Range[cell].Value = "\u25CB";
                ws.Range[cell].Font.Size = 15;                
                ws.Range[cell].RowHeight = height;
            }
        }
        public void Select(Microsoft.Office.Interop.Excel.Worksheet ws, int val)
        {
            Select(ws, GetByValue(val));
        }
        public void Select(Microsoft.Office.Interop.Excel.Worksheet ws, string rng)
        {
            ws.Range[rng].Value = "\u25CF"; //"\u25C9";
        }
        public string GetByValue(int value)
        {
            return SelPeers.First(kv => kv.Value == value).Key;
        }

        #endregion
    }
}
