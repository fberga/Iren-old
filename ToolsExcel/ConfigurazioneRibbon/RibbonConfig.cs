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

namespace ConfigurazioneRibbon
{
    public partial class RibbonConfig : Form
    {
        public RibbonConfig()
        {
            InitializeComponent();

            //setto ambiente di default (il primo scritto nel file di configurazione su Ambienti
            string[] ambienti = ConfigurationManager.AppSettings["Ambienti"].Split('|');
            ((CheckBox)groupBoxAmbienti.Controls["chkAmbiente" + ambienti[0]]).Checked = true;

            groupBoxAmbienti.Controls["chkAmbiente" + ambienti[0]].Click += AmbienteDefaultNnDisattivabile;

            //inizializzo connessione
            DataBase.InitNewDB(ambienti[0]);

            //carico la lista di utenti disponibili

            //carico la lista di applicazioni configurabili


        }

        private void AmbienteDefaultNnDisattivabile(object sender, EventArgs e)
        {
            ((CheckBox)sender).Checked = true;
        }

        private void CaricaListaUtenti()
        {
            //DataBase.Select();
        }

    }
}
