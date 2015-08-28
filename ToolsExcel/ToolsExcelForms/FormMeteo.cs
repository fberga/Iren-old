using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormMeteo : Form
    {
        DataView _entita;
        DataView _entitaProprieta;
        DateTime _dataRif;
        ACarica _carica;
        ARiepilogo _riepilogo;

        public FormMeteo(object dataRif, ACarica carica, ARiepilogo riepilogo)
        {
            InitializeComponent();

            _carica = carica;
            _entita = DataBase.LocalDB.Tables[DataBase.Tab.CATEGORIA_ENTITA].DefaultView;
            _entitaProprieta = DataBase.LocalDB.Tables[DataBase.Tab.ENTITA_PROPRIETA].DefaultView;
            _dataRif = (DateTime)dataRif;
            _riepilogo = riepilogo;

            labelData.Text = "Data Riferimento:   " + _dataRif.ToString("dddd dd MMMM yyyy");
            this.Text = Simboli.nomeApplicazione + " - Meteo";
        }

        private void frmMETEO_Load(object sender, EventArgs e)
        {
            _entitaProprieta.RowFilter = "SiglaProprieta = 'PROGR_IMPIANTO_TEMP_FONTE_ATTIVA' AND IdApplicazione = " + Simboli.AppID;

            string filtro = "";
            foreach (DataRowView prop in _entitaProprieta)
            {
                filtro += "'" + prop["SiglaEntita"] + "',";
            }

            if (filtro.Length > 0)
            {
                filtro = "SiglaEntita IN (" + filtro.Remove(filtro.Length - 1) + ")";
                _entita.RowFilter = filtro + " AND IdApplicazione = " + Simboli.AppID;
            }


            comboUP.DataSource = _entita;
            comboUP.DisplayMember = "DesEntita";

        }

        private void comboUP_SelectedIndexChanged(object sender, EventArgs e)
        {            
            if (DataBase.OpenConnection())
            {
                Array comboArray = groupDati.Controls.OfType<ComboBox>().ToArray();
                foreach (ComboBox cmb in comboArray)
                    groupDati.Controls.Remove(cmb);

                Array radioArray = groupDati.Controls.OfType<RadioButton>().ToArray();
                foreach (RadioButton rdb in radioArray)
                    groupDati.Controls.Remove(rdb);

                DataTable fonti = DataBase.Select(DataBase.SP.CHECK_FONTE_METEO, "@SiglaEntita=" + ((DataRowView)comboUP.SelectedItem)["SiglaEntita"] + ";@Data=" + _dataRif.ToString("yyyyMMdd")) ?? new DataTable();

                int fonteOrdine = 0;
                foreach (DataRow fonte in fonti.Rows)
                {
                    DateTime dataEmissione = DateTime.ParseExact(fonte["DataEmissione"].ToString(), "yyyyMMdd", CultureInfo.InvariantCulture);
                    if (!groupDati.Controls.ContainsKey("combo" + fonte["CodiceFonte"]))
                    {
                        ComboBox cmb = new ComboBox() 
                        { 
                            Name = "combo" + fonte["CodiceFonte"],
                            Font = groupDati.Font,
                            Location = new System.Drawing.Point(146, 50 + (28 * fonteOrdine) + 8),
                            Size = new System.Drawing.Size(190, 28),
                            FormattingEnabled = true
                        };
                        RadioButton rdb = new RadioButton()
                        {
                            Name = fonte["CodiceFonte"].ToString(),
                            Text = fonte["CodiceFonte"].ToString(),
                            Font = groupDati.Font,
                            Location = new System.Drawing.Point(5, 52 + (28 * fonteOrdine) + 8),
                            Size = new System.Drawing.Size(82, 24),
                            Checked = fonteOrdine == 0
                        };

                        rdb.CheckedChanged += rdb_CheckedChanged;

                        groupDati.Controls.Add(rdb);
                        groupDati.Controls.Add(cmb);
                        fonteOrdine++;
                    }
                    ((ComboBox)groupDati.Controls["combo" + fonte["CodiceFonte"]]).Items.Add(dataEmissione);
                    ((ComboBox)groupDati.Controls["combo" + fonte["CodiceFonte"]]).SelectedIndex = 0;
                }
            }
        }

        void rdb_CheckedChanged(object sender, EventArgs e)
        {
            if (DataBase.OpenConnection())
            {
                DataRowView entita = (DataRowView)comboUP.SelectedItem;
                RadioButton rbt = (RadioButton)sender;

                if (rbt.Checked)
                {
                    //TODO eliminare questo filtro e passare direttamente il codice della fonte (DA AGGIORNARE STRUTTURA SU DB)
                    _entitaProprieta.RowFilter = "SiglaProprieta = 'PROGR_IMPIANTO_TEMP_FONTE' AND SiglaEntita='" + entita["SiglaEntita"] + "' AND Valore = '" + rbt.Name + "' AND IdApplicazione = " + Simboli.AppID;

                    DataBase.Insert("spUpdateFonteMeteo", new Core.QryParams() 
                    {
                        {"@SiglaEntita", entita["SiglaEntita"]},
                        {"@Valore", _entitaProprieta[0]["Ordine"]}
                    });
                }
            }
            
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            if (DataBase.OpenConnection())
            {
                //TODO passare direttamente il codice della fonte (DA AGGIORNARE STRUTTURA SU DB)
                foreach (DataRowView entita in _entita)
                {
                    DataBase.Insert("spUpdateFonteMeteo", new Core.QryParams() 
                        {
                            {"@SiglaEntita", entita["SiglaEntita"]},
                            {"@Valore", "1"}
                        });
                }

                _entita.RowFilter = "IdApplicazione = " + Simboli.AppID;
                _entitaProprieta.RowFilter = "IdApplicazione = " + Simboli.AppID;
                this.Close();
            }
        }

        private void btnCarica_Click(object sender, EventArgs e)
        {
            btnCarica.Enabled = false;
            btnAnnulla.Enabled = false;

            object siglaEntita = ((DataRowView)comboUP.SelectedItem)["SiglaEntita"];

            string nomeCombo = "combo" + groupDati.Controls.OfType<RadioButton>().FirstOrDefault(btn => btn.Checked).Name;
            ComboBox cmb = groupDati.Controls.OfType<ComboBox>().FirstOrDefault(c => c.Name == nomeCombo);

            string dataEmissione = ((DateTime)cmb.SelectedItem).ToString("yyyyMMdd");

            bool gone = _carica.AzioneInformazione(siglaEntita, "METEO", "CARICA", _dataRif,  dataEmissione);

            _riepilogo.AggiornaRiepilogo(siglaEntita, "METEO", gone, _dataRif);

            Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogCarica, "Carica: Previsioni meteo");

            btnCarica.Enabled = true;
            btnAnnulla.Enabled = true;
        }
    }
}
