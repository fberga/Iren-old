using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System.Configuration;
using System.Text.RegularExpressions;

namespace ConfigurazioneRibbon
{
    public partial class RibbonConfig : Form
    {
        private int _maxCheckBox;
        private bool _abilitatoAllChecked = false;
        private bool _visibileAllChecked = false;
        private bool _statoAllChecked = false;

        private const string UTENTE_GRUPPO = "spUtenteGruppo";
        private const string APPLICAZIONI = "spApplicazioneProprieta";
        private const string RIBBON = "spApplicazioneRibbon";
        private const string UPDATE = "spUpdateApplicazioneRibbon";

        private DataTable _configurazioniDefault;
        private DataTable _configurazioneUtente;
        private string[] _ambienti;

        public RibbonConfig()
        {
            InitializeComponent();

            //setto ambiente di default (il primo scritto nel file di configurazione su Ambienti
            _ambienti = Workbook.AppSettings("Ambienti").Split('|');
            ((CheckBox)groupBoxAmbienti.Controls["chkAmbiente" + _ambienti[0]]).Checked = true;

            groupBoxAmbienti.Controls["chkAmbiente" + _ambienti[0]].Click += AmbienteDefaultNnDisattivabile;

            //inizializzo connessione
            DataBase.InitNewDB(_ambienti[0]);

            //carico i gruppi dal file di configurazione
            string[] gruppi = Workbook.AppSettings("Gruppi").Split('|');
            Dictionary<int, string> groupSource = new Dictionary<int,string>();

            foreach (string gruppo in gruppi)
            {
                string[] info = gruppo.Split(',');
                groupSource.Add(int.Parse(info[0]), info[1]);
            }

            cmbGruppi.ValueMember = "Key";
            cmbGruppi.DisplayMember = "Value";
            cmbGruppi.DataSource = groupSource.ToList();

            //carico la lista di applicazioni configurabili
            DataTable applicazioni = DataBase.Select(APPLICAZIONI, "@IdApplicazione=0");
            if (applicazioni != null)
            {
                cmbApplicazioni.DisplayMember = "DesApplicazione";
                cmbApplicazioni.ValueMember = "IdApplicazione";
                cmbApplicazioni.DataSource = applicazioni;
            }

            //trovo l'ultimo checkbox per identificare quanti componenti ci sono: il numero sara MaxCHK / 3
            var controls = GetAll(this, typeof(CheckBox));

            _maxCheckBox =
                (from ctrl in controls
                 where ctrl.Name.StartsWith("checkBox")
                 select int.Parse(Regex.Match(ctrl.Name, @"\d+").Value)).Max();
        }

        private void AmbienteDefaultNnDisattivabile(object sender, EventArgs e)
        {
            ((CheckBox)sender).Checked = true;
        }

        private void lbAbilitato_DoubleClick(object sender, EventArgs e)
        {
            _abilitatoAllChecked = !_abilitatoAllChecked;
            var controls = this.Controls.Cast<Control>();

            for (int i = 3; i <= _maxCheckBox; i += 3)
            {
                CheckBox chk = (CheckBox)this.Controls.Find("checkBox" + i, true).FirstOrDefault();
                chk.Checked = _abilitatoAllChecked;
            }
        }

        private void lbVisibile_DoubleClick(object sender, EventArgs e)
        {
            _visibileAllChecked = !_visibileAllChecked;
            var controls = this.Controls.Cast<Control>();

            for (int i = 2; i <= _maxCheckBox; i += 3)
            {
                CheckBox chk = (CheckBox)this.Controls.Find("checkBox" + i, true).FirstOrDefault();
                chk.Checked = _visibileAllChecked;
            }
        }

        private void lbStato_DoubleClick(object sender, EventArgs e)
        {
            _statoAllChecked = !_statoAllChecked;
            var controls = this.Controls.Cast<Control>();

            for (int i = 1; i <= _maxCheckBox; i += 3)
            {
                CheckBox chk = (CheckBox)this.Controls.Find("checkBox" + i, true).FirstOrDefault();
                chk.Checked = _statoAllChecked;
            }
        }

        public IEnumerable<Control> GetAll(Control control, Type type)
        {
            var controls = control.Controls.Cast<Control>();

            return controls.SelectMany(ctrl => GetAll(ctrl, type))
                                      .Concat(controls)
                                      .Where(c => c.GetType() == type);
        }

        private void cmbGruppi_SelectedIndexChanged(object sender, EventArgs e)
        {
            //carico la lista di utenti disponibili
            int id = ((KeyValuePair<int,string>)cmbGruppi.SelectedItem).Key;
            DataTable utenti = DataBase.Select(UTENTE_GRUPPO, "@IdUtenteGruppo=" + id);

            if (utenti != null)
            {
                listBoxUtenti.DataSource = utenti;
                listBoxUtenti.ValueMember = "IdUtente";
                listBoxUtenti.DisplayMember = "Nome";
                listBoxUtenti.SelectionMode = SelectionMode.MultiExtended;
            }
        }

        private void ApplyConfig(DataTable dt)
        {
            if (dt.Rows.Count == 0)
            {
                var controls = GetAll(this, typeof(CheckBox));

                foreach (CheckBox chk in controls)
                {
                    if (chk.Name.Contains("checkBox"))
                        chk.Checked = false;
                }
            }
            else
            {
                foreach (DataRow r in dt.Rows)
                {
                    if (!r["NomeControllo"].ToString().StartsWith("label"))
                    {
                        Control ctrl = this.Controls.Find(r["NomeControllo"].ToString(), true).First();

                        var chkBoxes = ctrl.Parent.Controls.OfType<CheckBox>().ToList();

                        chkBoxes[0].Checked = r["Stato"].Equals("1");
                        chkBoxes[1].Checked = r["Visibile"].Equals("1");
                        chkBoxes[2].Checked = r["Abilitato"].Equals("1");
                    }
                }
            }
        }

        private void btnApplyDefault_Click(object sender, EventArgs e)
        {
            ApplyConfig(_configurazioniDefault);
        }

        private void cmbApplicazioni_SelectedIndexChanged(object sender, EventArgs e)
        {
            //carico la tabella applicazioneRibbon default (cioè utente 62)       
            DataBase.DB.SetParameters("", 62, (int)cmbApplicazioni.SelectedValue);

            _configurazioniDefault = DataBase.Select(RIBBON) ?? new DataTable();

            if (listBoxUtenti.SelectedIndices.Count == 1)
            {
                DataBase.DB.SetParameters("", (int)listBoxUtenti.SelectedValue, (int)cmbApplicazioni.SelectedValue);
                _configurazioneUtente = DataBase.Select(RIBBON) ?? new DataTable();
                ApplyConfig(_configurazioneUtente);
            }
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            var ambienti = groupBoxAmbienti.Controls.OfType<CheckBox>().Where(chk => chk.Checked).ToList();

            foreach (CheckBox ambiente in ambienti)
            {
                if (ambiente.Name.Replace("chkAmbiente", "") != _ambienti[0])
                    DataBase.InitNewDB(ambiente.Name.Replace("chkAmbiente", ""));

                foreach (DataRowView user in listBoxUtenti.SelectedItems)
                {

                    DataBase.DB.SetParameters("", (int)user["IdUtente"], (int)cmbApplicazioni.SelectedValue);

                    foreach (DataRow r in _configurazioniDefault.Rows)
                    {
                        Control ctrl = this.Controls.Find(r["NomeControllo"].ToString(), true).First();

                        var chkBoxes = ctrl.Parent.Controls.OfType<CheckBox>().ToList();

                        DataBase.Insert(UPDATE, new Iren.ToolsExcel.Core.QryParams() 
                        {
                            {"@NomeControllo", ctrl.Name},
                            {"@Abilitato", chkBoxes[2].Checked ? "1" : "0"},
                            {"@Visibile", chkBoxes[1].Checked ? "1" : "0"},
                            {"@Stato", chkBoxes[0].Checked ? "1" : "0"}
                        });

                        //chkBoxes[0].Checked = r["Stato"].Equals("1");
                        //chkBoxes[1].Checked = r["Visibile"].Equals("1");
                        //chkBoxes[2].Checked = r["Abilitato"].Equals("1");
                    }
                }
            }
            DataBase.InitNewDB(_ambienti[0]);
        }

        private void listBoxUtenti_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxUtenti.SelectedIndices.Count == 1 && cmbApplicazioni.SelectedValue != null)
            {
                DataBase.DB.SetParameters("", (int)listBoxUtenti.SelectedValue, (int)cmbApplicazioni.SelectedValue);
                _configurazioneUtente = DataBase.Select(RIBBON) ?? new DataTable();
                ApplyConfig(_configurazioneUtente);
            }
        }
    }
}
