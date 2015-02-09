using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Iren.FrontOffice.Core;
using System.Globalization;
using Iren.FrontOffice.Base;

namespace Iren.FrontOffice.Forms
{
    public partial class frmMETEO : Form
    {
        DataView _entita;
        DataView _entitaProprieta;
        DateTime _dataRif;


        public frmMETEO(object dataRif)
        {
            InitializeComponent();

            _entita = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.CATEGORIAENTITA].DefaultView;
            _entitaProprieta = CommonFunctions.LocalDB.Tables[CommonFunctions.Tab.ENTITAPROPRIETA].DefaultView;
            _dataRif = (DateTime)dataRif;

            labelData.Text = "Data Riferimento: " + _dataRif.ToString("dd/MM/yyyy");
        }

        private void frmMETEO_Load(object sender, EventArgs e)
        {
            _entitaProprieta.RowFilter = "SiglaProprieta = 'PROGR_IMPIANTO_TEMP_FONTE_ATTIVA'";

            string filtro = "";
            foreach (DataRowView prop in _entitaProprieta)
            {
                filtro += "'" + prop["SiglaEntita"] + "',";
            }

            if (filtro.Length > 0)
            {
                filtro = "SiglaEntita IN (" + filtro.Remove(filtro.Length - 1) + ")";
                _entita.RowFilter = filtro;
            }


            comboUP.DataSource = _entita;
            comboUP.DisplayMember = "DesEntita";

        }

        private void comboUP_SelectedIndexChanged(object sender, EventArgs e)
        {            
            if (CommonFunctions.DB.OpenConnection())
            {
                Array comboArray = groupDati.Controls.OfType<ComboBox>().ToArray();
                foreach (ComboBox cmb in comboArray)
                    groupDati.Controls.Remove(cmb);

                Array radioArray = groupDati.Controls.OfType<RadioButton>().ToArray();
                foreach (RadioButton rdb in radioArray)
                    groupDati.Controls.Remove(rdb);

                DataView fonti = CommonFunctions.DB.Select("spCheckFonteMeteo", "@SiglaEntita=" + ((DataRowView)comboUP.SelectedItem)["SiglaEntita"] + ";@Data=" + _dataRif.ToString("yyyyMMdd")).DefaultView;

                int fonteOrdine = 0;
                foreach (DataRowView fonte in fonti)
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
                CommonFunctions.DB.CloseConnection();
            }
        }

        void rdb_CheckedChanged(object sender, EventArgs e)
        {
            if (CommonFunctions.DB.OpenConnection())
            {
                DataRowView entita = (DataRowView)comboUP.SelectedItem;
                RadioButton rbt = (RadioButton)sender;

                if (rbt.Checked)
                {
                    //TODO eliminare questo filtro e passare direttamente il codice della fonte (DA AGGIORNARE STRUTTURA SU DB)
                    _entitaProprieta.RowFilter = "SiglaProprieta = 'PROGR_IMPIANTO_TEMP_FONTE' AND SiglaEntita='" + entita["SiglaEntita"] + "' AND Valore = '" + rbt.Name + "'";

                    CommonFunctions.DB.Insert("spUpdateFonteMeteo", new QryParams() 
                    {
                        {"@SiglaEntita", entita["SiglaEntita"]},
                        {"@Valore", _entitaProprieta[0]["Ordine"]}
                    });
                }

                CommonFunctions.DB.CloseConnection();
            }
            
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            if (CommonFunctions.DB.OpenConnection())
            {
                //TODO passare direttamente il codice della fonte (DA AGGIORNARE STRUTTURA SU DB)
                foreach (DataRowView entita in _entita)
                {
                    CommonFunctions.DB.Insert("spUpdateFonteMeteo", new QryParams() 
                        {
                            {"@SiglaEntita", entita["SiglaEntita"]},
                            {"@Valore", "1"}
                        });
                }

                _entita.RowFilter = "";
                _entitaProprieta.RowFilter = "";
                CommonFunctions.DB.CloseConnection();
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

            bool gone = CommonFunctions.CaricaAzioneInformazione(siglaEntita, "METEO", "CARICA", _dataRif, dataEmissione);

            CommonFunctions.DB.OpenConnection();
            //TODO CommonFunctions.AggiornaRiepilogo(siglaEntita, METEO, gone)
            //TODO riabilitare log
            //CommonFunctions.InsertLog(DataBase.TipologiaLOG.LogCarica, "Carica: Previsioni meteo");
            //refresh true
            CommonFunctions.DB.CloseConnection();

            btnCarica.Enabled = true;
            btnAnnulla.Enabled = true;
        }
    }
}
