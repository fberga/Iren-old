using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Iren.ToolsExcel.Utility;
using Iren.ToolsExcel.Base;

namespace Iren.ToolsExcel.Forms
{
    public partial class ErrorPane : UserControl
    {
        public ErrorPane()
        {
            InitializeComponent();
        }

        private void ErrorPane_SizeChanged(object sender, EventArgs e)
        {
            Size sz = new Size(panelDescrizione.Width, int.MaxValue);
            sz = TextRenderer.MeasureText(lbTesto.Text, lbTesto.Font, sz, TextFormatFlags.WordBreak);

            panelDescrizione.Height = sz.Height + lbTitolo.Height + 15;
            panelContent.Height = this.Height - panelDescrizione.Height;
        }

        public void SetDimension(int width, int height)
        {
            this.Height = height;
            this.Width = width;
        }

        public void RefreshCheck(Check checkFunctions)
        {
            SplashScreen.UpdateStatus("Aggiornamento Check");
            foreach (Excel.Worksheet ws in Workbook.WB.Sheets)
            {
                NewDefinedNames newNomiDefiniti = new NewDefinedNames(ws.Name, NewDefinedNames.InitType.Check);
                if (newNomiDefiniti.HasCheck())
                {
                    foreach (CheckObj check in newNomiDefiniti.Checks)
                    {
                        TreeNode n = checkFunctions.ExecuteCheck(ws, newNomiDefiniti, check);

                        if (n.Nodes.Count > 0)
                        {
                            if(treeViewErrori.Nodes.ContainsKey(n.Name))
                            {
                                treeViewErrori.Nodes.RemoveByKey(n.Name);
                            }
                            treeViewErrori.Nodes.Add(n);
                        }
                    }
                }
            }
        }

        private void treeViewErrori_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Name.StartsWith("'"))
            {
                Excel.Range rng = (Excel.Range)Workbook.WB.Application.Range[e.Node.Name];
                rng.Worksheet.Activate();
                rng.Select();
            }
            else if (e.Node.Parent != null && e.Node.Parent.Name.StartsWith("'"))
            {
                Excel.Range rng = (Excel.Range)Workbook.WB.Application.Range[e.Node.Parent.Name];
                rng.Worksheet.Activate();
                rng.Select();
            }
        }
    }
}
