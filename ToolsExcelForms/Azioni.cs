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
        DataBase _db;


        public frmAZIONI(DataView categorie, DataView entita, DataView azioni, DataBase db)
        {
            InitializeComponent();
            _categorie = categorie;
            _entita = entita;
            _azioni = azioni;
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

        private void treeViewAzioni_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            e.Node.Checked = !e.Node.Checked;

            foreach (TreeNode node in e.Node.Nodes)
            {
                node.Checked = e.Node.Checked;
            }

        }

        private void treeViewAzioni_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = true;
        }

        private void treeViewAzioni_BeforeCheck(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = true;
            treeViewAzioni_NodeMouseClick(sender, e.Node);
        }
    }
}
