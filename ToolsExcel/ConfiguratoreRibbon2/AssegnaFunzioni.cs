using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.ToolsExcel.Utility;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    partial class AssegnaFunzioni : Form
    {
        public AssegnaFunzioni(IRibbonControl ctrl, RibbonGroup grp, int appID, int usrID)
        {
            InitializeComponent();

            DataTable ribbon = DataBase.Select(SP.APPLICAZIONE_UTENTE_RIBBON, "@IdApplicazione=" + appID + ";@IdUtente=" + usrID);
            if (ribbon != null)
            {
                var id =
                    (from r in ribbon.AsEnumerable()
                     where r["IdControllo"].Equals(ctrl.ID) && r["IdGruppo"].Equals(grp.ID)
                     select (int)r["IdGruppoControllo"]).FirstOrDefault();

                DataTable allFunctions = DataBase.Select(SP.FUNZIONE);
                DataTable ctrlFunctions = DataBase.Select(SP.CONTROLLO_FUNZIONE, "@IdGruppoControllo=" + id);

                if (ctrlFunctions != null)
                {
                    var usedFunctionsIds = new HashSet<object>(ctrlFunctions.AsEnumerable().Select(r => r["IdFunzione"]));
                    var unusedFunctions = allFunctions.AsEnumerable().Where(r => !usedFunctionsIds.Contains(r["IdFunzione"])).ToList();

                    foreach (var func in unusedFunctions)
                    {
                        TreeNode f = new TreeNode();
                        f.Text = func["NomeFunzione"].ToString();
                        f.Tag = func["IdFunzione"];

                        if (!treeViewNotUtilized.Nodes.ContainsKey(func["Evento"].ToString()))
                        {
                            TreeNode evento = new TreeNode(func["Evento"].ToString());
                            evento.Name = func["Evento"].ToString();
                            treeViewNotUtilized.Nodes.Add(evento);
                        }

                        treeViewNotUtilized.Nodes[func["Evento"].ToString()].Nodes.Add(f);
                    }

                }
            }
        }

        private void AggiungiFunzione_Click(object sender, EventArgs e)
        {
            TreeNode selected = treeViewNotUtilized.SelectedNode;
            if (selected.Tag != null)
            {
                string pName = selected.Parent.Name;
                if (!treeViewUtilized.Nodes.ContainsKey(pName))
                {
                    TreeNode evento = new TreeNode(pName);
                    evento.Name = pName;
                    treeViewUtilized.Nodes.Add(evento);
                }
                if (treeViewNotUtilized.Nodes[pName].Nodes.Count == 1)
                    treeViewNotUtilized.Nodes.Remove(selected.Parent);
                else
                    treeViewNotUtilized.Nodes[pName].Nodes.Remove(selected);

                treeViewUtilized.Nodes[pName].Nodes.Add(selected);
            }
        }

        private void RimuoviFunzione_Click(object sender, EventArgs e)
        {
            TreeNode selected = treeViewUtilized.SelectedNode;
            if (selected.Tag != null)
            {
                string pName = selected.Parent.Name;
                if (!treeViewNotUtilized.Nodes.ContainsKey(pName))
                {
                    TreeNode evento = new TreeNode(pName);
                    evento.Name = pName;
                    treeViewNotUtilized.Nodes.Add(evento);
                }
                if (treeViewUtilized.Nodes[pName].Nodes.Count == 1)
                    treeViewUtilized.Nodes.Remove(selected.Parent);
                else
                    treeViewUtilized.Nodes[pName].Nodes.Remove(selected);
                
                treeViewNotUtilized.Nodes[pName].Nodes.Add(selected);
            }
        }
    }
}
