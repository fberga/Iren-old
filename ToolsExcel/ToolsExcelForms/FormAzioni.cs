﻿using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormAzioni : Form
    {
        #region Variabili

        private DataView _azioni;
        private DataView _categorie;
        private DataView _categoriaEntita;
        private DataView _azioniCategorie;
        private DataView _entitaAzioni;
        private AEsporta _esporta;
        private ARiepilogo _r;
        private List<DateTime> _toProcessDates = new List<DateTime>();

        private FormSelezioneDate selDate = new FormSelezioneDate();

        private bool _giorniVisible = true;
        private bool _fromAzioni = false;

        #endregion

        #region Costruttori

        public FormAzioni(AEsporta esporta, ARiepilogo riepilogo)
        {
            InitializeComponent();

            _esporta = esporta;
            _r = riepilogo;

            _categorie = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA].DefaultView;
            _categorie.RowFilter = "";
            _categoriaEntita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            _categoriaEntita.RowFilter = "Gerarchia = '' OR Gerarchia IS NULL";
            _azioni = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE].DefaultView;
            _azioni.RowFilter = "Visibile = 1";
            _azioniCategorie = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE_CATEGORIA].DefaultView;
            _azioniCategorie.RowFilter = "";
            _entitaAzioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE].DefaultView;
            _entitaAzioni.RowFilter = "";
            DataView entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
            entitaProprieta.RowFilter = "";

            ConfigStructure();

            if (Struct.intervalloGiorni == 0 || !_giorniVisible)
            {
                comboGiorni.Text = DataBase.DataAttiva.ToString("dddd dd MMM yyyy");
                _toProcessDates.Add(DataBase.DataAttiva);
                comboGiorni.Enabled = false;
            }
            else
            {
                selDate.VisibleChanged += selDate_VisibleChanged;
            }
        }

        #endregion

        #region Metodi

        private void ConfigStructure()
        {
            System.Collections.IDictionary settings = (System.Collections.IDictionary)ConfigurationManager.GetSection("formSettings/azioniForm");

            Regex falseMatch = new Regex("false|0", RegexOptions.IgnoreCase);

            if (settings.Contains("CategorieVisible") && falseMatch.IsMatch(settings["CategorieVisible"].ToString()))
            {
                panelCategorie.Hide();
                Width -= panelCategorie.Width;
            }
            if (settings.Contains("GiorniVisible") && falseMatch.IsMatch(settings["GiorniVisible"].ToString()))
            {
                groupDate.Hide();
                _giorniVisible = false;
            }
            if (settings.Contains("GiorniVisible") && falseMatch.IsMatch(settings["GiorniVisible"].ToString()) &&
                settings.Contains("MercatiDaEsportareVisible") && falseMatch.IsMatch(settings["MercatiDaEsportareVisible"].ToString()))
            {
                panelTop.Hide();
                Height -= panelTop.Height;
            }
            if (settings.Contains("MeteoVisible") && falseMatch.IsMatch(settings["MeteoVisible"].ToString()))
            {
                btnMeteo.Hide();
            }
        }
        private void CaricaAzioni()
        {
            var stato = DataBase.DB.StatoDB;

            foreach (DataRowView azione in _azioni)
            {
                bool aggiungi = true;
                if (azione["IdFonte"] != DBNull.Value)
                {
                    var fonte = (Core.DataBase.NomiDB)Enum.Parse(typeof(Core.DataBase.NomiDB), azione["IdFonte"].ToString());
                    aggiungi = stato[fonte] == ConnectionState.Open;
                }

                if (azione["Operativa"].Equals("0") || (azione["Gerarchia"] is DBNull && aggiungi))
                {
                    treeViewAzioni.Nodes.Add(azione["SiglaAzione"].ToString(), azione["DesAzione"].ToString());
                }
                else if (aggiungi)
                {
                    if(treeViewAzioni.Nodes.ContainsKey(azione["Gerarchia"].ToString()))
                        treeViewAzioni.Nodes[azione["Gerarchia"].ToString()].Nodes.Add(azione["SiglaAzione"].ToString(), azione["DesAzione"].ToString());
                }
            }
            treeViewAzioni.ExpandAll();
        }
        private void CaricaCategorie()
        {
            foreach (DataRowView categoria in _categorie)
            {
                if (categoria["Operativa"].Equals("0"))
                {
                    treeViewCategorie.Nodes.Add(categoria["SiglaCategoria"].ToString(), categoria["DesCategoriaBreve"].ToString());
                }
                else
                {
                    if (categoria["Gerarchia"] is DBNull)
                        treeViewCategorie.Nodes.Add(categoria["SiglaCategoria"].ToString(), categoria["DesCategoriaBreve"].ToString());
                    else
                        treeViewCategorie.Nodes[categoria["Gerarchia"].ToString()].Nodes.Add(categoria["SiglaCategoria"].ToString(), categoria["DesCategoriaBreve"].ToString());
                }
            }
            treeViewCategorie.ExpandAll();
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
                if (!node.Checked)
                    notSel.Add(node.Name, false);
            }

            treeViewUP.Nodes.Clear();
            ThroughAllNodes(treeViewCategorie.Nodes, n =>
            {
                if (n.Checked)
                {
                    _categoriaEntita.RowFilter = "SiglaCategoria = '" + n.Name + "'";
                    foreach (DataRowView entita in _categoriaEntita)
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
        private void Evidenzia(TreeNode node, bool evidenzia)
        {
            if (evidenzia)
            {
                node.BackColor = System.Drawing.Color.Gold;
                node.ForeColor = System.Drawing.Color.DarkRed;
            }
            else
            {
                node.BackColor = treeViewAzioni.BackColor;
                node.ForeColor = treeViewAzioni.ForeColor;
                node.NodeFont = treeViewAzioni.Font;
            }
        }
        private void ColoraNodi()
        {
            ThroughAllNodes(treeViewAzioni.Nodes, n =>
            {
                Evidenzia(n, n.Checked);
            });
            ThroughAllNodes(treeViewCategorie.Nodes, n =>
            {
                Evidenzia(n, n.Checked);
            });
            ThroughAllNodes(treeViewUP.Nodes, n =>
            {
                Evidenzia(n, n.Checked);
            });
        }
        private void CheckParents()
        {
            foreach (TreeNode node in treeViewAzioni.Nodes)
            {
                if (node.Nodes.Count != 0)
                {
                    if (HasCheckedNode(node))
                        node.Checked = true;
                    else
                        node.Checked = false;
                }
            }
            foreach (TreeNode node in treeViewCategorie.Nodes)
            {
                if (node.Nodes.Count != 0)
                {
                    if (HasCheckedNode(node))
                        node.Checked = true;
                    else
                        node.Checked = false;
                }
            }
        }

        #endregion


        private void SelectParents(TreeNode node, Boolean isChecked)
        {
            var parent = node.Parent;

            if (parent == null)
                return;

            if (!isChecked && HasCheckedNode(parent))
                return;

            parent.Checked = isChecked;
            SelectParents(parent, isChecked);
        }
        private bool HasCheckedNode(TreeNode node)
        {
            return node.Nodes.Cast<TreeNode>().Any(n => n.Checked);
        }

        #region Eventi

        private void frmAZIONI_Load(object sender, EventArgs e)
        {
            this.Text = Simboli.nomeApplicazione + " - Azioni";
            CaricaAzioni();
            CaricaCategorie();
        }
        private void treeView_BeforeCollapse(object sender, TreeViewCancelEventArgs e)
        {
            e.Cancel = true;
        }
        //private void treeViewAzioni_AfterCheck(object sender, TreeViewEventArgs e)
        //{
        //    //treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;

        //    //if (e.Node.Nodes.Count > 0)
        //    //    foreach (TreeNode node in e.Node.Nodes)
        //    //        node.Checked = e.Node.Checked;
        //    //else
        //    //    SelectParents(e.Node, e.Node.Checked);

        //    //treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;

        //    //treeViewCategorie.AfterCheck -= treeViewCategorie_AfterCheck;

        //    //ThroughAllNodes(treeViewCategorie.Nodes, nodoCategoria =>
        //    //{
        //    //    nodoCategoria.Checked = false;
        //    //});

        //    //ThroughAllNodes(treeViewAzioni.Nodes, nodoAzione =>
        //    //{
        //    //    if (nodoAzione.Nodes.Count == 0)
        //    //    {
        //    //        _azioniCategorie.RowFilter = "SiglaAzione = '" + nodoAzione.Name + "'";
        //    //        foreach (DataRowView azioneCategoria in _azioniCategorie)
        //    //        {
        //    //            TreeNode nodoCategoria = treeViewCategorie.Nodes.Find(azioneCategoria["SiglaCategoria"].ToString(), true)[0];
        //    //            nodoCategoria.Checked = nodoCategoria.Checked || nodoAzione.Checked;
        //    //        }
        //    //    }
        //    //});

        //    //treeViewCategorie.AfterCheck += treeViewCategorie_AfterCheck;


        //    //_fromAzioni = true;

        //    //if (e.Node.Checked)
        //    //{
        //    //    if (e.Node.Nodes.Count == 0)
        //    //    {
        //    //        treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;
        //    //        if (e.Node.Parent != null && !e.Node.Parent.Checked)
        //    //            e.Node.Parent.Checked = true;
        //    //        treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;

        //    //        _azioniCategorie.RowFilter = "SiglaAzione = '" + e.Node.Name + "'";
                    
        //    //        foreach (DataRowView azioneCategoria in _azioniCategorie)
        //    //            if (!treeViewCategorie.Nodes.Find(azioneCategoria["SiglaCategoria"].ToString(), true)[0].Checked)
        //    //                treeViewCategorie.Nodes.Find(azioneCategoria["SiglaCategoria"].ToString(), true)[0].Checked = true;
        //    //    }
        //    //    else
        //    //    {
        //    //        foreach (TreeNode node in e.Node.Nodes)
        //    //            if(!node.Checked)
        //    //                node.Checked = true;
        //    //    }
        //    //}
        //    //else
        //    //{
        //    //    if (e.Node.Nodes.Count > 0)
        //    //    {
        //    //        foreach (TreeNode node in e.Node.Nodes)
        //    //            if (node.Checked)
        //    //                node.Checked = false;
        //    //    }
        //    //    else
        //    //    {
        //    //        treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;
        //    //        if(e.Node.Parent.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
        //    //            e.Node.Parent.Checked = false;
        //    //        treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;
        //    //    }

        //    //    Dictionary<string, bool> cateogorie = new Dictionary<string, bool>();
        //    //    ThroughAllNodes(treeViewCategorie.Nodes, n => 
        //    //    {
        //    //        if (n.Nodes.Count == 0)
        //    //        {
        //    //            cateogorie.Add(n.Name, false);
        //    //        }
        //    //    });

        //    //    ThroughAllNodes(treeViewAzioni.Nodes, n =>
        //    //    {
        //    //        if (n.Nodes.Count == 0 && n.Checked)
        //    //        {
        //    //            _azioniCategorie.RowFilter = "SiglaAzione = '" + n.Name + "'";
        //    //            foreach (DataRowView azioneCategoria in _azioniCategorie)
        //    //            {
        //    //                cateogorie[azioneCategoria["SiglaCategoria"].ToString()] = true;
        //    //            }
        //    //        }
        //    //    });

        //    //    treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;
        //    //    foreach (KeyValuePair<string, bool> cat in cateogorie)
        //    //    {
        //    //        if (!cat.Value && treeViewCategorie.Nodes.Find(cat.Key, true)[0].Checked)
        //    //            treeViewCategorie.Nodes.Find(cat.Key, true)[0].Checked = false;
        //    //    }
        //    //    treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;
        //    //}

        //    ////CaricaEntita();
        //    ////ColoraNodi();
        //    //_fromAzioni = false;
        //}

        private void treeView_AfterCheck(object sender, TreeViewEventArgs e)
        {
            TreeView from = sender as TreeView;
            TreeView to = from.Name == "treeViewAzioni" ? treeViewCategorie : treeViewAzioni;

            from.AfterCheck -= treeView_AfterCheck;
            to.AfterCheck -= treeView_AfterCheck;

            if (e.Node.Nodes.Count > 0)
                foreach (TreeNode node in e.Node.Nodes)
                    node.Checked = e.Node.Checked;

            string filter = from.Name == "treeViewAzioni" ? "SiglaAzione" : "SiglaCategoria";
            string field = from.Name == "treeViewAzioni" ? "SiglaCategoria" : "SiglaAzione";

            Dictionary<string, bool> checkedNodes = new Dictionary<string, bool>();
            if (e.Node.Checked)
            {
                ThroughAllNodes(from.Nodes, n =>
                {
                    if (n.Nodes.Count == 0 && n.Checked)
                    {
                        _azioniCategorie.RowFilter = filter + " = '" + n.Name + "'";
                        foreach (DataRowView azioneCategoria in _azioniCategorie)
                        {
                            TreeNode n1 = to.Nodes.Find(azioneCategoria[field].ToString(), true)[0];
                            if (checkedNodes.ContainsKey(azioneCategoria[field].ToString()))
                                checkedNodes[azioneCategoria[field].ToString()] = true;
                            else
                                checkedNodes.Add(azioneCategoria[field].ToString(), true);
                        }
                    }
                });

                ThroughAllNodes(to.Nodes, n =>
                {
                    if (n.Nodes.Count == 0 && checkedNodes.ContainsKey(n.Name))
                        n.Checked = n.Checked || checkedNodes[n.Name];
                });
            }
            else
            {
                ThroughAllNodes(from.Nodes, n =>
                {
                    if (n.Nodes.Count == 0)
                    {
                        _azioniCategorie.RowFilter = filter + " = '" + n.Name + "'";
                        foreach (DataRowView azioneCategoria in _azioniCategorie)
                        {
                            TreeNode n1 = to.Nodes.Find(azioneCategoria[field].ToString(), true)[0];
                            if (checkedNodes.ContainsKey(azioneCategoria[field].ToString()))
                                checkedNodes[azioneCategoria[field].ToString()] = checkedNodes[azioneCategoria[field].ToString()] || n.Checked;
                            else
                                checkedNodes.Add(azioneCategoria[field].ToString(), n.Checked);
                        }
                    }
                });

                ThroughAllNodes(to.Nodes, n =>
                {
                    if (n.Nodes.Count == 0)
                        n.Checked = n.Checked && checkedNodes[n.Name];
                });
            }

            CheckParents();
            CaricaEntita();
            ColoraNodi();

            to.AfterCheck += treeView_AfterCheck;
            from.AfterCheck += treeView_AfterCheck;
        }

        //private void treeViewCategorie_AfterCheck(object sender, TreeViewEventArgs e)
        //{
        //    treeViewCategorie.AfterCheck -= treeViewCategorie_AfterCheck;

        //    if (e.Node.Nodes.Count > 0)
        //        foreach (TreeNode node in e.Node.Nodes)
        //            node.Checked = e.Node.Checked;
        //    else
        //        SelectParents(e.Node, e.Node.Checked);

        //    treeViewCategorie.AfterCheck += treeViewCategorie_AfterCheck;

        //    treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;
            
        //    if (e.Node.Checked)
        //    {
        //        ThroughAllNodes(treeViewAzioni.Nodes, nodoAzione =>
        //        {
        //            nodoAzione.Checked = false;
        //        });

        //        ThroughAllNodes(treeViewCategorie.Nodes, nodoCategoria =>
        //        {
        //            if (nodoCategoria.Nodes.Count == 0 && nodoCategoria.Checked)
        //            {
        //                _azioniCategorie.RowFilter = "SiglaCategoria = '" + nodoCategoria.Name + "'";
        //                foreach (DataRowView azioneCategoria in _azioniCategorie)
        //                {
        //                    TreeNode nodoAzione = treeViewAzioni.Nodes.Find(azioneCategoria["SiglaAzione"].ToString(), true)[0];
        //                    nodoAzione.Checked = nodoCategoria.Checked;
        //                }
        //            }
        //        });
        //    }
        //    else
        //    {
        //        Dictionary<string, bool> checkedNodes = new Dictionary<string, bool>();
        //        ThroughAllNodes(treeViewCategorie.Nodes, nodoCategoria =>
        //        {
        //            if (nodoCategoria.Nodes.Count == 0)
        //            {
        //                _azioniCategorie.RowFilter = "SiglaCategoria = '" + nodoCategoria.Name + "'";
        //                foreach (DataRowView azioneCategoria in _azioniCategorie)
        //                {
        //                    TreeNode nodoAzione = treeViewAzioni.Nodes.Find(azioneCategoria["SiglaAzione"].ToString(), true)[0];
        //                    if (checkedNodes.ContainsKey(azioneCategoria["SiglaAzione"].ToString()))
        //                        checkedNodes[azioneCategoria["SiglaAzione"].ToString()] = checkedNodes[azioneCategoria["SiglaAzione"].ToString()] || nodoCategoria.Checked;
        //                    else
        //                        checkedNodes.Add(azioneCategoria["SiglaAzione"].ToString(), nodoCategoria.Checked);
        //                }
        //            }
        //        });

        //        ThroughAllNodes(treeViewAzioni.Nodes, nodoAzione =>
        //        {
        //            if (nodoAzione.Nodes.Count == 0)
        //                nodoAzione.Checked = nodoAzione.Checked && checkedNodes[nodoAzione.Name];
        //        });
        //    }
            
        //    treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;


        //    //treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;

        //    ////ThroughAllNodes(treeViewAzioni.Nodes, nodoAzione =>
        //    ////{
        //    ////    nodoAzione.Checked = false;
        //    ////});

        //    //Dictionary<string, bool> checkedNodes = new Dictionary<string, bool>();

        //    //ThroughAllNodes(treeViewCategorie.Nodes, nodoCategoria =>
        //    //{
        //    //    if (nodoCategoria.Nodes.Count == 0)
        //    //    {
        //    //        _azioniCategorie.RowFilter = "SiglaCategoria = '" + nodoCategoria.Name + "'";
        //    //        foreach (DataRowView azioneCategoria in _azioniCategorie)
        //    //        {
        //    //            TreeNode nodoAzione = treeViewAzioni.Nodes.Find(azioneCategoria["SiglaAzione"].ToString(), true)[0];
        //    //            if (checkedNodes.ContainsKey(azioneCategoria["SiglaAzione"].ToString()))
        //    //                checkedNodes[azioneCategoria["SiglaAzione"].ToString()] = checkedNodes[azioneCategoria["SiglaAzione"].ToString()] || nodoCategoria.Checked;
        //    //            else
        //    //                checkedNodes.Add(azioneCategoria["SiglaAzione"].ToString(), nodoCategoria.Checked);
        //    //        }
        //    //    }
        //    //});

        //    //ThroughAllNodes(treeViewAzioni.Nodes, nodoAzione =>
        //    //{
        //    //    if (nodoAzione.Nodes.Count == 0)
        //    //        nodoAzione.Checked = nodoAzione.Checked && checkedNodes[nodoAzione.Name];
        //    //});

        //    //treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;



        //    //if (e.Node.Checked)
        //    //{
        //    //    if (e.Node.Nodes.Count > 0)
        //    //    {
        //    //        foreach (TreeNode node in e.Node.Nodes)
        //    //            if (!node.Checked)
        //    //                node.Checked = true;
        //    //    }
        //    //    else
        //    //    {
        //    //        treeViewCategorie.AfterCheck -= treeViewCategorie_AfterCheck;
        //    //        if (e.Node.Parent != null && !e.Node.Parent.Checked)
        //    //            e.Node.Parent.Checked = true;
        //    //        treeViewCategorie.AfterCheck += treeViewCategorie_AfterCheck;
        //    //    }

        //    //    if (!_fromAzioni)
        //    //    {
        //    //        treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;
        //    //        _azioniCategorie.RowFilter = "SiglaCategoria = '" + e.Node.Name + "'";
        //    //        foreach (DataRowView azioneCategoria in _azioniCategorie)
        //    //            if (!treeViewAzioni.Nodes.Find(azioneCategoria["SiglaAzione"].ToString(), true)[0].Checked)
        //    //                treeViewAzioni.Nodes.Find(azioneCategoria["SiglaAzione"].ToString(), true)[0].Checked = true;
        //    //        CheckParents(treeViewAzioni);
        //    //        treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;
        //    //    }
        //    //}
        //    //else
        //    //{
        //    //    if (e.Node.Nodes.Count > 0)
        //    //    {
        //    //        foreach (TreeNode node in e.Node.Nodes)
        //    //            if (node.Checked)
        //    //                node.Checked = false;
        //    //    }
        //    //    else
        //    //    {
        //    //        treeViewCategorie.AfterCheck -= treeViewCategorie_AfterCheck;
        //    //        if (e.Node.Parent != null && e.Node.Parent.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
        //    //            e.Node.Parent.Checked = false;
        //    //        treeViewCategorie.AfterCheck += treeViewCategorie_AfterCheck;
        //    //    }
        //    //    List<string> azioni = new List<string>();
        //    //    bool trovato = false;

        //    //    ThroughAllNodes(treeViewAzioni.Nodes, n =>
        //    //    {
        //    //        if (n.Nodes.Count == 0)
        //    //        {
        //    //            _azioniCategorie.RowFilter = "SiglaAzione = '" + n.Name + "'";
        //    //            foreach (DataRowView azioneCategoria in _azioniCategorie)
        //    //                trovato = trovato || treeViewCategorie.Nodes.Find(azioneCategoria["SiglaCategoria"].ToString(), true)[0].Checked;

        //    //            if (!trovato && n.Checked)
        //    //                n.Checked = false;
        //    //        }
        //    //    });

        //    //    CheckParents(treeViewAzioni);
        //    //}
        //    //CaricaEntita();
        //    //ColoraNodi();
        //}
        private void treeViewUP_AfterCheck(object sender, TreeViewEventArgs e)
        {
            checkTutte.CheckedChanged -= checkTutte_CheckedChanged;
            
            Evidenzia(e.Node, e.Node.Checked);

            bool check = true;
            foreach (TreeNode node in treeViewUP.Nodes)
            {
                check = check && node.Checked;
            }
            checkTutte.Checked = check;

            checkTutte.CheckedChanged += checkTutte_CheckedChanged;
        }
        private void btnMeteo_Click(object sender, EventArgs e)
        {
            if (_toProcessDates.Count > 0)
            {
                if ((_toProcessDates.Count > 1 && MessageBox.Show("Ci sono più date selezionate. Procedere con la prima?", Simboli.nomeApplicazione, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes) || _toProcessDates.Count == 1)
                {
                    FormMeteo meteo = new FormMeteo(_toProcessDates[0]);
                    meteo.ShowDialog();
                }
            }
            else
                MessageBox.Show("Non è stata selezionata alcuna data...", Simboli.nomeApplicazione, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        private void btnApplica_Click(object sender, EventArgs e)
        {
            btnApplica.Enabled = false;
            btnAnnulla.Enabled = false;
            btnMeteo.Enabled = false;

            if (_toProcessDates.Count == 0)
                MessageBox.Show("Non è stata selezionata alcuna data...", Simboli.nomeApplicazione, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else if (treeViewUP.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
                MessageBox.Show("Non è stata selezionata alcuna unità...", Simboli.nomeApplicazione, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                SplashScreen.Show();

                foreach (DateTime date in _toProcessDates)
                {
                    bool calcola = false;
                    int count = 0;
                    ThroughAllNodes(treeViewAzioni.Nodes, n =>
                    {
                        if (n.Nodes.Count == 0 && n.Checked)
                        {
                            ThroughAllNodes(treeViewUP.Nodes, n1 =>
                            {
                                if (n1.Checked)
                                    count += (n1.Name == "UP_BUS" ? 2 : 1);
                            });
                            if (n.Parent.Name == "CARICA")
                                calcola = true;
                        }
                    });

                    if (calcola)
                    {
                        ThroughAllNodes(treeViewUP.Nodes, n =>
                        {
                            if (n.Checked)
                                count++;
                        });
                    }

                    if (DataBase.OpenConnection())
                    {
                        string suffissoData = Date.GetSuffissoData(date);

                        bool[] statoAzione = new bool[4] { false, false, false, false };

                        ThroughAllNodes(treeViewAzioni.Nodes, nodoAzione =>
                        {
                            if (nodoAzione.Checked && nodoAzione.Nodes.Count == 0)
                            {
                                TreeNode[] nodiEntita = treeViewUP.Nodes.OfType<TreeNode>().Where(node => node.Checked).ToArray();
                                _azioni.RowFilter = "SiglaAzione = '" + nodoAzione.Name + "'";

                                ThroughAllNodes(treeViewUP.Nodes, nodoEntita =>
                                {
                                    if (nodoEntita.Checked && nodoEntita.Nodes.Count == 0)
                                    {
                                        string nomeFoglio = NewDefinedNames.GetSheetName(nodoEntita.Name);
                                        bool presente;

                                        SplashScreen.UpdateStatus(nodoAzione.Parent.Text + " " + nodoAzione.Text + ": " + nodoEntita.Text);

                                        switch (nodoAzione.Parent.Name)
                                        {
                                            case "CARICA":
                                                presente = Workbook.CaricaAzioneInformazione(nodoEntita.Name, nodoAzione.Name, nodoAzione.Parent.Name, date);
                                                _r.AggiornaRiepilogo(nodoEntita.Name, nodoAzione.Name, presente, date);
                                                statoAzione[0] = true;
                                                break;
                                            case "GENERA":
                                                presente = Workbook.CaricaAzioneInformazione(nodoEntita.Name, nodoAzione.Name, nodoAzione.Parent.Name, date);
                                                _r.AggiornaRiepilogo(nodoEntita.Name, nodoAzione.Name, presente, date);
                                                statoAzione[1] = true;
                                                break;
                                            case "ESPORTA":
                                                presente = _esporta.RunExport(nodoEntita.Name, nodoAzione.Name, nodoEntita.Text, nodoAzione.Text, date);
                                                _r.AggiornaRiepilogo(nodoEntita.Name, nodoAzione.Name, presente, date);
                                                statoAzione[2] = true;
                                                break;
                                        }

                                        if (_azioni[0]["Relazione"] != DBNull.Value && Struct.visualizzaRiepilogo)
                                        {
                                            string[] azioneRelazione = _azioni[0]["Relazione"].ToString().Split(';');
                                            
                                            NewDefinedNames newNomiDefiniti = new NewDefinedNames("Main");
                                            Excel.Worksheet ws = Workbook.WB.Sheets["Main"];

                                            foreach (string relazione in azioneRelazione)
                                            {
                                                _azioni.RowFilter = "SiglaAzione = '" + relazione + "'";

                                                Range rng = new Range(newNomiDefiniti.GetRowByName(nodoEntita.Name), newNomiDefiniti.GetColFromName(relazione, suffissoData));
                                                if (ws.Range[rng.ToString()].Interior.ColorIndex != 2)
                                                {
                                                    ws.Range[rng.ToString()].Value = "RI" + _azioni[0]["Gerarchia"];
                                                    Style.RangeStyle(ws.Range[rng.ToString()], fontSize: 8, bold: true, foreColor: 3, backColor: 6, align: Excel.XlHAlign.xlHAlignCenter);
                                                }
                                            }
                                            _azioni.RowFilter = "SiglaAzione = '" + nodoAzione.Name + "'";
                                        }
                                    }
                                });
                                switch (nodoAzione.Parent.Name)
                                {
                                    case "CARICA":
                                        Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogCarica, "Carica: " + nodoAzione.Name);
                                        break;
                                    case "GENERA":
                                        Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogGenera, "Genera: " + nodoAzione.Name);
                                        break;
                                    case "ESPORTA":
                                        Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogEsporta, "Esporta: " + nodoAzione.Name);
                                        break;
                                }
                            }
                        });

                        if (statoAzione[0] || statoAzione[1])
                        {
                            Sheet.SalvaModifiche();
                            DataBase.SalvaModificheDB();
                        }

                        DataBase.DB.CloseConnection();
                    }
                }

                SplashScreen.Close();
            }

            btnApplica.Enabled = true;
            btnAnnulla.Enabled = true;
            btnMeteo.Enabled = true;
        }
        private void checkTutte_CheckedChanged(object sender, EventArgs e)
        {
            treeViewUP.AfterCheck -= treeViewUP_AfterCheck;
            foreach (TreeNode node in treeViewUP.Nodes)
            {
                node.Checked = checkTutte.Checked;
                Evidenzia(node, checkTutte.Checked);
            }
            treeViewUP.AfterCheck += treeViewUP_AfterCheck;
        }
        private void FormAzioni_FormClosing(object sender, FormClosingEventArgs e)
        {
            selDate.Close();
        }
        private void comboGiorni_MouseClick(object sender, EventArgs e)
        {
            selDate.BringToFront();
            selDate.Top = comboGiorni.PointToScreen(Point.Empty).Y + comboGiorni.Height;
            selDate.Left = comboGiorni.PointToScreen(Point.Empty).X;
            selDate.Width = comboGiorni.Width;
            selDate.Show();
        }
        private void comboGiorni_TextChanged(object sender, EventArgs e)
        {
            comboGiorni.TextChanged -= comboGiorni_TextChanged;
            if (comboGiorni.Text == "" || comboGiorni.Text == "- Click per selezionare le date -")
            {
                comboGiorni.Text = "- Click per selezionare le date -";
                comboGiorni.Font = new Font(comboGiorni.Font, FontStyle.Italic);
                comboGiorni.ForeColor = SystemColors.ControlDarkDark;
            }
            else
            {
                comboGiorni.Font = new Font(comboGiorni.Font, FontStyle.Regular);
                comboGiorni.ForeColor = SystemColors.ControlText;
            }
            comboGiorni.TextChanged += comboGiorni_TextChanged;
        }
        private void selDate_VisibleChanged(object sender, EventArgs e)
        {
            if (!_toProcessDates.SequenceEqual(selDate.SelectedDates))
            {
                _toProcessDates = new List<DateTime>(selDate.SelectedDates);

                comboGiorni.Text = "";
                if (_toProcessDates.Count == 1)
                    comboGiorni.Text = _toProcessDates[0].ToString("ddd dd MMM yyyy");
                else if (_toProcessDates.Count > 0)
                    comboGiorni.Text = _toProcessDates.Count + " giorni";
                else
                    comboGiorni.Text = "";
            }
        }

        #endregion

    }
}
