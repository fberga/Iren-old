using Iren.PSO.Base;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Forms
{
    public partial class FormImportXML : Form
    {
        private DataTable _tabellaImportXML;
        private List<CommonInfo> _commonInfo;

        public FormImportXML()
        {
            this.Text = Simboli.NomeApplicazione + " - Import XML (Emergenza)";

            InitializeComponent();
            btnApri_Click(null, null);
        }

        private void openFileXMLImport_FileOk(object sender, CancelEventArgs e)
        {
            DataSet tmp = new DataSet();
            tmp.ReadXml(openFileXMLImport.FileName);
            _tabellaImportXML = tmp.Tables[DataBase.TAB.EXPORT_XML];

            foreach(DataColumn c in Workbook.Repository[DataBase.TAB.EXPORT_XML].Columns) 
            {
                if(!_tabellaImportXML.Columns.Contains(c.ColumnName)) 
                {
                    System.Windows.Forms.MessageBox.Show("Il file selezionato non deriva da un export. Non è possibile importare informazioni da questo file!", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    e.Cancel = true;
                    return;
                }
            }

            //controllo date
            var strDataMin =
                (from r in _tabellaImportXML.AsEnumerable()
                 select r["Data"].ToString().Substring(0, 8)).Min();

            var strDataMax =
                (from r in _tabellaImportXML.AsEnumerable()
                 select r["Data"].ToString().Substring(0, 8)).Max();

            DateTime dataMin = DateTime.ParseExact(strDataMin, "yyyyMMdd", CultureInfo.InvariantCulture);
            DateTime dataMax = DateTime.ParseExact(strDataMax, "yyyyMMdd", CultureInfo.InvariantCulture);

            if(Workbook.DataAttiva < dataMin.Date || Workbook.DataAttiva > dataMax.Date)
            {
                System.Windows.Forms.MessageBox.Show("Il file non contiene elementi con date compatibili a quelle del foglio aperto. Non è possibile importare informazioni da questo file!", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Cancel = true;
                return;
            }

            var nomeApplicazione =
                (from r in Workbook.Repository[DataBase.TAB.LISTA_APPLICAZIONI].AsEnumerable()
                 where r["IdApplicazione"].Equals(int.Parse(_tabellaImportXML.Rows[0]["IdApplicazione"].ToString()))
                 select r["DesApplicazione"]).First();

            //segnalo da che applicativo è stato creato l'XML
            richTextInfoTop.Text ="XML generato da: " + nomeApplicazione + "\n\nLista delle informazioni recuperate dal file XML che sono compatibili con il foglio corrente:";

            //tabella in ordine, posso procedere con la visualizzazione dei campi coinvolti
            SetTreeView();
        }


        private void SetTreeView()
        {
            //creo tabella con corrispondenza dei campi tra XML e foglio excel
            //Entità - Informazione - SiglaEntità - SiglaInformazione
            var dtImport =
                (from r in _tabellaImportXML.AsEnumerable()
                 select new { SiglaEntita = r["SiglaEntita"], SiglaInformazione = r["SiglaInformazione"] }).Distinct();

            _commonInfo = new List<CommonInfo>();

            foreach (var ele in dtImport)
            {
                var info =
                    (from r in Workbook.Repository[DataBase.TAB.ENTITA_INFORMAZIONE].AsEnumerable()
                     where r["SiglaEntita"].Equals(ele.SiglaEntita) && r["SiglaInformazione"].Equals(ele.SiglaInformazione)
                     select r).FirstOrDefault();

                if (info != null)
                {
                    var desEntita =
                        (from r in Workbook.Repository[DataBase.TAB.CATEGORIA_ENTITA].AsEnumerable()
                         where r["SiglaEntita"].Equals(info["SiglaEntita"])
                         select r["DesEntita"]).First();

                    CommonInfo c = new CommonInfo(info["SiglaEntita"], desEntita);

                    if (_commonInfo.Contains(c))
                        _commonInfo[_commonInfo.IndexOf(c)].Info.Add(info["SiglaInformazione"], info["DesInformazione"]);
                    else
                    {
                        c.Info.Add(info["SiglaInformazione"], info["DesInformazione"]);
                        _commonInfo.Add(c);
                    }
                }
            }

            if (_commonInfo.Count == 0)
            {
                btnImporta.Enabled = false;
            }
            else
            {
                btnImporta.Enabled = true;
                foreach (var c in _commonInfo)
                {
                    tvEntitaInformazioni.Nodes.Add(c.SiglaEntita, c.DesEntita);
                    foreach (var kv in c.Info)
                        tvEntitaInformazioni.Nodes[c.SiglaEntita].Nodes.Add(kv.Key.ToString(), kv.Value.ToString());
                }
            }
        }

        private void btnApri_Click(object sender, EventArgs e)
        {
            var directory = Workbook.Repository.Applicazione["PathDatiComuniEmergenza"].ToString();
            if (Directory.Exists(directory))
                openFileXMLImport.InitialDirectory = directory;

            openFileXMLImport.Filter = "XML Files (*.xml)|*.xml";
            openFileXMLImport.ShowDialog();
        }

        private void btnImporta_Click(object sender, EventArgs e)
        {
            SplashScreen.Show();
            Workbook.ScreenUpdating = false;
            Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationManual;
            Sheet.Protected = false;

            foreach (var c in _commonInfo)
            {
                SplashScreen.UpdateStatus("Importo dati per " + c.DesEntita);
                string foglio = DefinedNames.GetSheetName(c.SiglaEntita);
                Excel.Worksheet ws = Workbook.Sheets[foglio];
                DefinedNames definedNames = new DefinedNames(foglio);

                foreach (var kv in c.Info)
                {
                    var values =
                        from r in _tabellaImportXML.AsEnumerable()
                        where r["SiglaEntita"].Equals(c.SiglaEntita) && r["SiglaInformazione"].Equals(kv.Key) &&
                            (r["Data"].ToString().Substring(0,8).CompareTo(Workbook.DataAttiva.ToString("yyyyMMdd")) >= 0)
                        select new { Data = r["Data"], Valore = r["Valore"] };

                    foreach (var val in values)
                    {
                        string suffissoData = Date.GetSuffissoData(val.Data.ToString());
                        string suffissoOra = Date.GetSuffissoOra(val.Data);

                        Range rng = new Range();
                        if (definedNames.TryGet(out rng, c.SiglaEntita, kv.Key, suffissoData, suffissoOra))
                        {
                            ws.Range[rng.ToString()].Value = val.Valore;
                        }
                    }
                }
            }
            
            Sheet.Protected = true;
            Workbook.Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            SplashScreen.Close();
            Workbook.ScreenUpdating = true;
        }
    }

    public class CommonInfo : IEquatable<CommonInfo>
    {
        private object _siglaEntita;
        private object _desEntita;
        private Dictionary<object, object> _info = new Dictionary<object, object>();


        public string SiglaEntita { get { return _siglaEntita.ToString(); } }
        public string DesEntita { get { return _desEntita.ToString(); } }
        public Dictionary<object, object> Info { get { return _info; } }


        public CommonInfo(object siglaEntita, object desEntita)
        {
            _siglaEntita = siglaEntita;
            _desEntita = desEntita;
        }

        public bool Equals(CommonInfo other)
        {
            return other.SiglaEntita.Equals(_siglaEntita);
        }
    }



}
