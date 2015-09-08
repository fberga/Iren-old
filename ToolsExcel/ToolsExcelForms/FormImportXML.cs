using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Iren.ToolsExcel.Forms
{
    public partial class FormImportXML : Form
    {
        private DataTable _tabellaImportXML;

        public FormImportXML()
        {
            InitializeComponent();
            btnApri_Click(null, null);
        }

        private void openFileXMLImport_FileOk(object sender, CancelEventArgs e)
        {
            DataSet tmp = new DataSet();
            tmp.ReadXml(openFileXMLImport.FileName);
            _tabellaImportXML = tmp.Tables[DataBase.Tab.EXPORT_XML];

            foreach(DataColumn c in DataBase.LocalDB.Tables[DataBase.Tab.EXPORT_XML].Columns) 
            {
                if(!_tabellaImportXML.Columns.Contains(c.ColumnName)) 
                {
                    //TODO segnalare all'utente
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

            if(DataBase.DataAttiva < dataMin.Date || DataBase.DataAttiva > dataMax.Date)
            {
                //TODO segnalare all'utente
                e.Cancel = true;
                return;
            }

            //tabella in ordine, posso procedere con la visualizzazione dei campi coinvolti
            SetDataGridViews();
        }


        private void SetDataGridViews()
        {
            //creo tabella con corrispondenza dei campi tra XML e foglio excel
            //Entità - Informazione - SiglaEntità - SiglaInformazione
            DataTable dt = _tabellaImportXML.DefaultView.ToTable(true, "SiglaEntita", "SiglaInformazione");

            dataGridFileXML.DataSource = dt;
        }

        private void btnApri_Click(object sender, EventArgs e)
        {
            var path = Workbook.GetUsrConfigElement("emergenza");
            string cartellaEmergenza = Esporta.PreparePath(path.Value);

            openFileXMLImport.InitialDirectory = cartellaEmergenza;
            openFileXMLImport.ShowDialog();
        }
    }
}
