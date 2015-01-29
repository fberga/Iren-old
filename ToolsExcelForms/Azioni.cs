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
        DataView _entitaAzioni;
        DataBase _db;


        public frmAZIONI(DataView categorie, DataView entita, DataView azioni, DataView azioniCategorie, DataView entitaAzioni, DataBase db)
        {
            InitializeComponent();
            _categorie = categorie;
            _entita = entita;
            _azioni = azioni;
            _azioniCategorie = azioniCategorie;
            _entitaAzioni = entitaAzioni;
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
            //var stato = _db.StatoDB();

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

        private void ThroughAllNodes(TreeNodeCollection root, Action<TreeNode> callback)
        {
            if (root.Count > 0)
            {
                foreach (TreeNode node in root.OfType<TreeNode>())
                {
                    callback(node);
                    ThroughAllNodes(node.Nodes, callback);
                }
            }
        }

        private void CaricaEntita()
        {
            Dictionary<string, bool> notSel = new Dictionary<string, bool>();

            foreach (TreeNode node in treeViewUP.Nodes)
            {
                if(!node.Checked)
                notSel.Add(node.Name, false);
            }

            treeViewUP.Nodes.Clear();
            ThroughAllNodes(treeViewCategorie.Nodes, n =>
            {
                if (n.Checked)
                {
                    _entita.RowFilter = "SiglaCategoria = '" + n.Name + "'";
                    foreach (DataRowView entita in _entita)
                    {
                        ThroughAllNodes(treeViewAzioni.Nodes, n1 =>
                        {
                            if (n1.Checked)
                            {
                                _entitaAzioni.RowFilter = "SiglaEntita = '" + entita["SiglaEntita"] + "' AND SiglaAzione = '" + n1.Name + "'";
                                if (_entitaAzioni.Count > 0 && treeViewUP.Nodes.Find(entita["SiglaEntita"].ToString(), true).Length == 0)
                                {
                                    treeViewUP.Nodes.Add(entita["SiglaEntita"].ToString(), entita["DesEntita"].ToString());
                                    if (notSel.ContainsKey(entita["SiglaEntita"].ToString()))
                                        treeViewUP.Nodes[entita["SiglaEntita"].ToString()].Checked = false;
                                    else
                                        treeViewUP.Nodes[entita["SiglaEntita"].ToString()].Checked = true;
                                }
                            }
                        });
                    }
                }
            });
        }

        private void CheckParents()
        {
            foreach (TreeNode node in treeViewAzioni.Nodes)
            {
                if (node.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length > 0)
                    if(!node.Checked)
                        node.Checked = true;
            }
            foreach (TreeNode node in treeViewCategorie.Nodes)
            {
                if (node.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length > 0)
                    if (!node.Checked)
                        node.Checked = true;
            }
        }

        bool fromAzioni = false;
        private void treeViewAzioni_AfterCheck1(object sender, TreeViewEventArgs e)
        {
            fromAzioni = true;

            if (e.Node.Checked)
            {
                if (e.Node.Nodes.Count == 0)
                {
                    treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck1;
                    if (e.Node.Parent != null && !e.Node.Parent.Checked)
                        e.Node.Parent.Checked = true;
                    treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck1;

                    _azioniCategorie.RowFilter = "SiglaAzione = '" + e.Node.Name + "'";
                    
                    foreach (DataRowView azioneCategoria in _azioniCategorie)
                        if (!treeViewCategorie.Nodes.Find(azioneCategoria["SiglaCategoria"].ToString(), true)[0].Checked)
                            treeViewCategorie.Nodes.Find(azioneCategoria["SiglaCategoria"].ToString(), true)[0].Checked = true;
                }
                else
                {
                    foreach (TreeNode node in e.Node.Nodes)
                        if(!node.Checked)
                            node.Checked = true;
                }
            }
            else
            {
                if (e.Node.Nodes.Count > 0)
                {
                    foreach (TreeNode node in e.Node.Nodes)
                        if (node.Checked)
                            node.Checked = false;
                }
                else
                {
                    treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck1;
                    if(e.Node.Parent.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
                        e.Node.Parent.Checked = false;
                    treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck1;
                }

                Dictionary<string, bool> cateogorie = new Dictionary<string, bool>();
                ThroughAllNodes(treeViewCategorie.Nodes, n => 
                {
                    if (n.Nodes.Count == 0)
                    {
                        cateogorie.Add(n.Name, false);
                    }
                });

                ThroughAllNodes(treeViewAzioni.Nodes, n =>
                {
                    if (n.Nodes.Count == 0 && n.Checked)
                    {
                        _azioniCategorie.RowFilter = "SiglaAzione = '" + n.Name + "'";
                        foreach (DataRowView azioneCategoria in _azioniCategorie)
                        {
                            cateogorie[azioneCategoria["SiglaCategoria"].ToString()] = true;
                        }
                    }
                });
                
                treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck1;
                foreach (KeyValuePair<string, bool> cat in cateogorie)
                {
                    if (!cat.Value && treeViewCategorie.Nodes.Find(cat.Key, true)[0].Checked)
                        treeViewCategorie.Nodes.Find(cat.Key, true)[0].Checked = false;
                }
                treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck1;
            }

            //CheckParents();
            CaricaEntita();
            fromAzioni = false;
        }

        private void treeViewCategorie_AfterCheck1(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Checked)
            {
                if (e.Node.Nodes.Count > 0)
                {
                    foreach (TreeNode node in e.Node.Nodes)
                        if (!node.Checked)
                            node.Checked = true;
                }
                else
                {
                    treeViewCategorie.AfterCheck -= treeViewCategorie_AfterCheck1;
                    if (e.Node.Parent != null && !e.Node.Parent.Checked)
                        e.Node.Parent.Checked = true;
                    treeViewCategorie.AfterCheck += treeViewCategorie_AfterCheck1;
                }

                if (!fromAzioni)
                {
                    _azioniCategorie.RowFilter = "SiglaCategoria = '" + e.Node.Name + "'";
                    foreach (DataRowView azioneCategoria in _azioniCategorie)
                        if (!treeViewAzioni.Nodes.Find(azioneCategoria["SiglaAzione"].ToString(), true)[0].Checked)
                            treeViewAzioni.Nodes.Find(azioneCategoria["SiglaAzione"].ToString(), true)[0].Checked = true;
                }
            }
            else
            {

                if (e.Node.Nodes.Count > 0)
                {
                    foreach (TreeNode node in e.Node.Nodes)
                        if (node.Checked)
                            node.Checked = false;
                }
                else
                {
                    treeViewCategorie.AfterCheck -= treeViewCategorie_AfterCheck1;
                    if (e.Node.Parent.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
                        e.Node.Parent.Checked = false;
                    treeViewCategorie.AfterCheck += treeViewCategorie_AfterCheck1;
                }
                List<string> azioni = new List<string>();
                bool trovato = false;

                treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck1;

                ThroughAllNodes(treeViewAzioni.Nodes, n =>
                {
                    if (n.Nodes.Count == 0)
                    {
                        _azioniCategorie.RowFilter = "SiglaAzione = '" + n.Name + "'";
                        foreach (DataRowView azioneCategoria in _azioniCategorie)
                            trovato = trovato || treeViewCategorie.Nodes.Find(azioneCategoria["SiglaCategoria"].ToString(), true)[0].Checked;

                        if (!trovato && !n.Checked)
                            n.Checked = false;
                    }
                });

                treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck1;
            }

            //CheckParents();
            CaricaEntita();
        }
    }
}
