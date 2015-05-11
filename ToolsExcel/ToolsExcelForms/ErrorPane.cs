﻿using System;
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
        public static Font GetFont { get { return new ErrorPane().treeViewErrori.Font; } }

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
            NewDefinedNames gotos = new NewDefinedNames("Main", NewDefinedNames.InitType.GOTOsOnly);

            //Reset delle celle GOTO di tutto il Workbook
            List<string> gotoRanges = gotos.GetAllFromAddressGOTO();
            foreach (string gotoCell in gotoRanges)
                Style.RangeStyle(Workbook.WB.Application.Range[gotoCell], backColor: 2, foreColor: 1);

            foreach (Excel.Worksheet ws in Workbook.WB.Sheets)
            {
                if(ws.Name != "Main" && ws.Name != "Log") 
                {
                    NewDefinedNames newNomiDefiniti = new NewDefinedNames(ws.Name, NewDefinedNames.InitType.CheckNaming);

                    if (newNomiDefiniti.HasCheck())
                    {
                        foreach (CheckObj check in newNomiDefiniti.Checks)
                        {
                            CheckOutput o = checkFunctions.ExecuteCheck(ws, newNomiDefiniti, check);

                            if (o.Node.Nodes.Count > 0)
                            {
                                if (treeViewErrori.Nodes.ContainsKey(o.Node.Name))
                                {
                                    treeViewErrori.Nodes.RemoveByKey(o.Node.Name);
                                }
                                treeViewErrori.Nodes.Add(o.Node);


                                //Coloro le celle GOTO del Main e della scheda corrente
                                List<string> rngToCheck = gotos.GetFromAddressGOTO(check.SiglaEntita);
                                foreach (string gotoCell in rngToCheck)
                                {
                                    if(o.Status == CheckOutput.CheckStatus.Error)
                                        Style.RangeStyle(Workbook.WB.Application.Range[gotoCell], backColor: 3, foreColor: 6);
                                    else if(o.Status == CheckOutput.CheckStatus.Alert)
                                        Style.RangeStyle(Workbook.WB.Application.Range[gotoCell], backColor: 6, foreColor: 3);
                                }

                                //Coloro la barra del titolo verticale
                                Range titoloVert = new Range(check.Range);
                                //riduco il range ad una sola cella alla colonna B
                                titoloVert.StartColumn = 2;
                                titoloVert.ColOffset = 1;
                                titoloVert.RowOffset = 1;
                                if(o.Status == CheckOutput.CheckStatus.Error)
                                    Style.RangeStyle(ws.Range[titoloVert.ToString()].MergeArea, backColor: 3, foreColor: 6);
                                else if(o.Status == CheckOutput.CheckStatus.Alert)
                                    Style.RangeStyle(ws.Range[titoloVert.ToString()].MergeArea, backColor: 6, foreColor: 3);
                            }
                            else
                            {
                                //Reset della barra del titolo verticale
                                Range titoloVert = new Range(check.Range);
                                //riduco il range ad una sola cella alla colonna B
                                titoloVert.StartColumn = 2;
                                titoloVert.ColOffset = 1;
                                titoloVert.RowOffset = 1;
                                Style.RangeStyle(ws.Range[titoloVert.ToString()].MergeArea, backColor: 2, foreColor: 1);
                            }
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

        public void SelectNode(string address)
        {
            TreeNode[] nodes = treeViewErrori.Nodes.Find(address, true);
            treeViewErrori.CollapseAll();
            if (nodes.Length > 0)
            {                
                nodes[0].Expand();
                TreeNode n = nodes[0].Parent;
                while (n != null)
                {
                    n.Expand();
                    n = n.Parent;
                }

                treeViewErrori.SelectedNode = nodes[0];
                //treeViewErrori.Focus();
            }
        }     
    }
}
