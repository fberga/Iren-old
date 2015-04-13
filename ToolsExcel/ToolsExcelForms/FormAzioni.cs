using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormAzioni : Form
    {
        #region Variabili

        DataView _azioni;
        DataView _categorie;
        DataView _categoriaEntita;
        DataView _azioniCategorie;
        DataView _entitaAzioni;
        DataView _entitaProprieta;
        AEsporta _esporta;
        ARiepilogo _r;

        //bool _categorieVisible = true;
        //bool _mercatiDaEsportareVisible = true;
        //bool _meteoVisible = true;
        bool _giorniVisible = true;

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
            _categoriaEntita.RowFilter = "";
            _azioni = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE].DefaultView;
            _azioni.RowFilter = "Visibile = 1";
            _azioniCategorie = DataBase.LocalDB.Tables[DataBase.Tab.AZIONE_CATEGORIA].DefaultView;
            _azioniCategorie.RowFilter = "";
            _entitaAzioni = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_AZIONE].DefaultView;
            _entitaAzioni.RowFilter = "";
            _entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
            _entitaProprieta.RowFilter = "";

            System.Data.DataTable dt = new System.Data.DataTable()
            {
                Columns =
                {
                    {"DescData", typeof(string)},
                    {"Data", typeof(DateTime)}
                }
            };

            for (int i = 0; i <= Struct.intervalloGiorni; i++ )
            {
                DataRow r = dt.NewRow();
                r["DescData"] = (i + 1) + "° - " + DataBase.DB.DataAttiva.AddDays(i).ToString("dd/MM/yyyy");
                r["Data"] = DataBase.DB.DataAttiva.AddDays(i);

                dt.Rows.Add(r);
            }
            comboGiorni.DataSource = dt;
            comboGiorni.DisplayMember = "DescData";            

            if (comboGiorni.Items.Count == 1)
            {
                comboGiorni.SelectedIndex = 0;
                comboGiorni.Enabled = false;
            }
            else
            {
                comboGiorni.SelectedIndex = -1;
            }

            ConfigStructure();
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
                //_categorieVisible = false;
            }
            if (settings.Contains("GiorniVisible") && falseMatch.IsMatch(settings["GiorniVisible"].ToString()))
            {
                groupDate.Hide();
                _giorniVisible = false;
            }
            if (settings.Contains("MercatiDaEsportareVisible") && falseMatch.IsMatch(settings["MercatiDaEsportareVisible"].ToString()))
            {
                groupMercati.Hide();
                //_mercatiDaEsportareVisible = false;
            }
            if (settings.Contains("MeteoVisible") && falseMatch.IsMatch(settings["MeteoVisible"].ToString()))
            {
                groupMercati.Hide();
                //_meteoVisible = false;
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

        private void CheckParents(TreeView treeView)
        {
            foreach (TreeNode node in treeView.Nodes)
            {
                if (node.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
                {
                    if (node.Checked)
                        node.Checked = false;
                }
                else
                {
                    if (!node.Checked)
                        node.Checked = true;
                }
            }
        }

        #endregion

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

        bool fromAzioni = false;
        private void treeViewAzioni_AfterCheck(object sender, TreeViewEventArgs e)
        {
            fromAzioni = true;

            if (e.Node.Checked)
            {
                if (e.Node.Nodes.Count == 0)
                {
                    treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;
                    if (e.Node.Parent != null && !e.Node.Parent.Checked)
                        e.Node.Parent.Checked = true;
                    treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;

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
                    treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;
                    if(e.Node.Parent.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
                        e.Node.Parent.Checked = false;
                    treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;
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

                treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;
                foreach (KeyValuePair<string, bool> cat in cateogorie)
                {
                    if (!cat.Value && treeViewCategorie.Nodes.Find(cat.Key, true)[0].Checked)
                        treeViewCategorie.Nodes.Find(cat.Key, true)[0].Checked = false;
                }
                treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;
            }

            CaricaEntita();
            ColoraNodi();
            fromAzioni = false;
        }

        private void treeViewCategorie_AfterCheck(object sender, TreeViewEventArgs e)
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
                    treeViewCategorie.AfterCheck -= treeViewCategorie_AfterCheck;
                    if (e.Node.Parent != null && !e.Node.Parent.Checked)
                        e.Node.Parent.Checked = true;
                    treeViewCategorie.AfterCheck += treeViewCategorie_AfterCheck;
                }

                if (!fromAzioni)
                {
                    treeViewAzioni.AfterCheck -= treeViewAzioni_AfterCheck;
                    _azioniCategorie.RowFilter = "SiglaCategoria = '" + e.Node.Name + "'";
                    foreach (DataRowView azioneCategoria in _azioniCategorie)
                        if (!treeViewAzioni.Nodes.Find(azioneCategoria["SiglaAzione"].ToString(), true)[0].Checked)
                            treeViewAzioni.Nodes.Find(azioneCategoria["SiglaAzione"].ToString(), true)[0].Checked = true;
                    CheckParents(treeViewAzioni);
                    treeViewAzioni.AfterCheck += treeViewAzioni_AfterCheck;
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
                    treeViewCategorie.AfterCheck -= treeViewCategorie_AfterCheck;
                    if (e.Node.Parent.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
                        e.Node.Parent.Checked = false;
                    treeViewCategorie.AfterCheck += treeViewCategorie_AfterCheck;
                }
                List<string> azioni = new List<string>();
                bool trovato = false;

                ThroughAllNodes(treeViewAzioni.Nodes, n =>
                {
                    if (n.Nodes.Count == 0)
                    {
                        _azioniCategorie.RowFilter = "SiglaAzione = '" + n.Name + "'";
                        foreach (DataRowView azioneCategoria in _azioniCategorie)
                            trovato = trovato || treeViewCategorie.Nodes.Find(azioneCategoria["SiglaCategoria"].ToString(), true)[0].Checked;

                        if (!trovato && n.Checked)
                            n.Checked = false;
                    }
                });

                CheckParents(treeViewAzioni);
            }
            CaricaEntita();
            ColoraNodi();
        }

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
            if (comboGiorni.SelectedIndex != -1)
            {
                FormMeteo meteo = new FormMeteo(((DataRowView)comboGiorni.SelectedItem)["Data"]);
                meteo.ShowDialog();
            }
            else
                MessageBox.Show("Non è stata selezionata alcuna data...", Simboli.nomeApplicazione, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void btnApplica_Click(object sender, EventArgs e)
        {
            btnApplica.Enabled = false;
            btnAnnulla.Enabled = false;
            btnMeteo.Enabled = false;

            if (_giorniVisible && comboGiorni.SelectedIndex == -1)
                MessageBox.Show("Non è stata selezionata alcuna data...", Simboli.nomeApplicazione, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else if (treeViewUP.Nodes.OfType<TreeNode>().Where(n => n.Checked).ToArray().Length == 0)
                MessageBox.Show("Non è stata selezionata alcuna unità...", Simboli.nomeApplicazione, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
            {
                
                DateTime dataRif;
                if (_giorniVisible)
                    dataRif = (DateTime)((DataRowView)comboGiorni.SelectedItem)["Data"];
                else
                    dataRif = DataBase.DataAttiva;

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
                    string suffissoData = Date.GetSuffissoData(DataBase.DB.DataAttiva, dataRif);

                    bool[] statoAzione = new bool[4] { false, false, false, false };

                    ThroughAllNodes(treeViewAzioni.Nodes, n =>
                    {
                        if (n.Checked && n.Nodes.Count == 0)
                        {
                            TreeNode[] nodiEntita = treeViewUP.Nodes.OfType<TreeNode>().Where(node => node.Checked).ToArray();
                            _azioni.RowFilter = "SiglaAzione = '" + n.Name + "'";

                            ThroughAllNodes(treeViewUP.Nodes, n1 =>
                            {
                                if (n1.Checked && n1.Nodes.Count == 0)
                                {
                                    string nomeFoglio = DefinedNames.GetSheetName(n1.Name);
                                    bool presente;
                                    switch (n.Parent.Name)
                                    {
                                        case "CARICA":
                                            presente = Workbook.CaricaAzioneInformazione(n1.Name, n.Name, n.Parent.Name, dataRif);
                                            _r.AggiornaRiepilogo(n1.Name, n.Name, presente);
                                            statoAzione[0] = true;
                                            break;
                                        case "GENERA":
                                            presente = Workbook.CaricaAzioneInformazione(n1.Name, n.Name, n.Parent.Name, dataRif);
                                            _r.AggiornaRiepilogo(n1.Name, n.Name, presente);
                                            statoAzione[1] = true;
                                            break;
                                        case "ESPORTA":
                                            presente = _esporta.RunExport(n1.Name, n.Name, n1.Text, n.Text, dataRif);
                                            _r.AggiornaRiepilogo(n1.Name, n.Name, presente);
                                            statoAzione[2] = true;
                                            break;
                                    }

                                    if (_azioni[0]["Relazione"] != DBNull.Value)
                                    {
                                        string[] azioneRelazione = _azioni[0]["Relazione"].ToString().Split(';');
                                        foreach (string relazione in azioneRelazione)
                                        {
                                            _azioni.RowFilter = "SiglaAzione = '" + relazione + "'";

                                            if (DefinedNames.IsDefined("Main", DefinedNames.GetName("RIEPILOGO", n1.Name, relazione, suffissoData)))
                                            {
                                                DefinedNames nomiDefiniti = new DefinedNames("Main");
                                                Tuple<int, int> cella = nomiDefiniti[DefinedNames.GetName("RIEPILOGO", n1.Name, relazione, suffissoData)][0];

                                                Excel.Worksheet ws = Workbook.WB.Sheets["Main"];
                                                if (ws.Cells[cella.Item1, cella.Item2].Interior.ColoIndex != 2)
                                                {
                                                    ws.Cells[cella.Item1, cella.Item2].Value = "RI" + _azioni[0]["Gerarchia"];
                                                    Style.RangeStyle(ws.Cells[cella.Item1, cella.Item2], "Bold:True;ForeColor:3;BackColor:6;Align:Center");
                                                }
                                            }
                                        }
                                        _azioni.RowFilter = "SiglaAzione = '" + n.Name + "'";
                                    }
                                }
                            });
                            switch (n.Parent.Name)
                            {
                                case "CARICA":
                                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogCarica, "Carica: " + n.Name);
                                    break;
                                case "GENERA":
                                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogGenera, "Genera: " + n.Name);
                                    break;
                                case "ESPORTA":
                                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogEsporta, "Esporta: " + n.Name);
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

            btnApplica.Enabled = true;
            btnAnnulla.Enabled = true;
            btnMeteo.Enabled = true;
        }

        #endregion

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
    }
}
