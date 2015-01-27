using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.FrontOffice.Core;

namespace Iren.FrontOffice.Forms
{
    public partial class frmAZIONI : Form
    {
        DataView _azioni;
        DataView _categorie;
        DataView _entita;
        DataView _azioniCategorie;
        DataBase _db;


        public frmAZIONI(DataView categorie, DataView entita, DataView azioni, DataView azioniCategorie, DataBase db)
        {
            InitializeComponent();
            _categorie = categorie;
            _entita = entita;
            _azioni = azioni;
            _azioniCategorie = azioniCategorie;
            _db = db;
        }

        private void CaricaAzioni()
        {
            var stato = _db.StatoDB();

            foreach (DataRowView azione in _azioni)
            {
                bool aggiungi = true;
                if (azione["IdFonte"] != DBNull.Value)
                {
                    var fonte = (DataBase.NomiDB)Enum.Parse(typeof(DataBase.NomiDB), azione["IdFonte"].ToString());
                    aggiungi = stato[fonte] == ConnectionState.Open;
                }

                if (azione["Operativa"].Equals("0") || (azione["Gerarchia"] is DBNull && aggiungi))
                {
                    treeViewAzioni.Nodes.Add(azione["SiglaAzione"].ToString(), azione["DesAzione"].ToString());
                }
                else if (aggiungi)
                {
                    treeViewAzioni.Nodes[azione["Gerarchia"].ToString()].Nodes.Add(azione["SiglaAzione"].ToString(), azione["DesAzione"].ToString());
                }
            }
            treeViewAzioni.ExpandAll();
        }

        private void CaricaCategorie()
        {
            var stato = _db.StatoDB();

            foreach (DataRowView categoria in _categorie)
            {
                if (categoria["Operativa"].Equals("0"))
                {
                    treeViewCategorie.Nodes.Add(categoria["SiglaCategoria"].ToString(), categoria["DesCategoriaBreve"].ToString());
                }
                else
                {
                    treeViewCategorie.Nodes[categoria["Gerarchia"].ToString()].Nodes.Add(categoria["SiglaCategoria"].ToString(), categoria["DesCategoriaBreve"].ToString());
                }
            }
            treeViewCategorie.ExpandAll();
        }

        private void frmAZIONI_Load(object sender, EventArgs e)
        {
            CaricaAzioni();
            CaricaCategorie();
        }
        
        private void treeView_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = true;
        }

        private void ThroughAllNode(TreeNodeCollection root, Action<TreeNode> callback)
        {
            if (root.Count > 0)
            {
                foreach (TreeNode node in root.OfType<TreeNode>())
                {
                    callback(node);
                    ThroughAllNode(node.Nodes, callback);
                }
            }
        }

        private void treeView_AfterCheck(object sender, TreeViewEventArgs e)
        {
            TreeView sourceTreeView = (TreeView)sender;
            TreeView destTreeView = sourceTreeView;
            bool check = e.Node.Checked;
            if (e.Node.Nodes.Count == 0)
            {
                string filter = "";
                string key = "";
                switch (sourceTreeView.Name)
                {
                    case "treeViewAzioni":
                        filter = "SiglaAzione";
                        key = "SiglaCategoria";
                        destTreeView = (TreeView)Controls.Find("treeViewCategorie", true)[0];
                        break;
                    case "treeViewCategorie":
                        filter = "SiglaCategoria";
                        key = "SiglaAzione";
                        destTreeView = (TreeView)Controls.Find("treeViewAzioni", true)[0];
                        break;
                }
                //disabilito la callback in caso di evento check per evitare loop
                destTreeView.AfterCheck -= treeView_AfterCheck;

                //modifico stato del padre
                sourceTreeView.AfterCheck -= treeView_AfterCheck;
                if (e.Node.Parent.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0) 
                {
                    e.Node.Parent.Checked = false;
                    e.Node.Parent.BackColor = sourceTreeView.BackColor;
                }
                else
                {
                    e.Node.Parent.Checked = true;
                    e.Node.Parent.BackColor = Color.Coral;
                }       
                sourceTreeView.AfterCheck += treeView_AfterCheck;
                
                //elimino tutti i check
                ThroughAllNode(destTreeView.Nodes, n => n.Checked = false);
                
                //ripristino quelli necessari
                ThroughAllNode(sourceTreeView.Nodes, n => 
                {
                    if (n.Checked)
                    {
                        _azioniCategorie.RowFilter = filter + " = '" + n.Name + "'";
                        foreach (DataRowView azioneCategoria in _azioniCategorie)
                        {
                            destTreeView.Nodes.Find(azioneCategoria[key].ToString(), true)[0].Checked = true;
                            destTreeView.Nodes.Find(azioneCategoria[key].ToString(), true)[0].Parent.Checked = true;
                        }
                    }
                });
                destTreeView.AfterCheck += treeView_AfterCheck;
            }
            else
            {
                foreach (TreeNode node in e.Node.Nodes)
                {
                    node.Checked = check;
                }
            }
        }        
    }
}
