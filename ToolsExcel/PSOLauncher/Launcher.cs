using Iren.ToolsExcel.Utility;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PSOLauncher
{
    public partial class Launcher : Form
    {
        const string CONTROLLO_APPLICAZIONE = "RIBBON.spControlloApplicazione";

        private ImageList _listaImmaginiApplicazioni = new ImageList();

        public int IdUtente { get; private set; }

        public Launcher()
        {
            InitializeComponent();

            var resourceSet = Iren.ToolsExcel.Base.Properties.Resources.ResourceManager.GetResourceSet(CultureInfo.InstalledUICulture, true, true);
            var imgs =
                from r in resourceSet.Cast<DictionaryEntry>()
                where r.Value is Image
                select r;

            foreach (var img in imgs)
                _listaImmaginiApplicazioni.Images.Add(img.Key as string, img.Value as Image);

            _listaImmaginiApplicazioni.ImageSize = new System.Drawing.Size(32, 32);

            DataBase.CreateNew(ConfigurationManager.AppSettings["DB"]);

            if (!InitUsr())
            {
                Close();
            }

            DataBase.IdUtente = IdUtente;            
            
            if (!InitApplicazioni())
            {
                Close();
            }

        }

        private bool InitUsr()
        {
            DataTable dtUtente = DataBase.Select(DataBase.SP.UTENTE, "@CodUtenteWindows="+Environment.UserName);

            if (dtUtente != null && dtUtente.Rows.Count > 0)
            {
                IdUtente = (int)dtUtente.Rows[0]["IdUtente"];
                return true;
            }

            return false;
        }

        private bool InitApplicazioni()
        {
            DataTable dtApplicazioni = DataBase.Select(DataBase.SP.APPLICAZIONE, "@IdApplicazione=0");
            if (dtApplicazioni != null && dtApplicazioni.Rows.Count > 0)
            {
                int j = 0;
                for (int i = 0; i < dtApplicazioni.Rows.Count; i++)
                {
                    DataTable dtControllo = DataBase.Select(CONTROLLO_APPLICAZIONE, "@IdApplicazione=" + dtApplicazioni.Rows[i]["IdApplicazione"]);

                    if(dtControllo != null && dtControllo.Rows.Count > 0) 
                    {
                        if (!menuIconTray.Items.ContainsKey(dtControllo.Rows[0]["Nome"].ToString()))
                        {
                            Button btn = new Button();
                            btn.ImageList = _listaImmaginiApplicazioni;
                            btn.ImageKey = dtControllo.Rows[0]["Immagine"].ToString();
                            btn.Text = dtControllo.Rows[0]["Label"].ToString();
                            btn.TextImageRelation = TextImageRelation.ImageBeforeText;
                            btn.Name = dtControllo.Rows[0]["Nome"].ToString();
                            btn.Tag = dtControllo.Rows[0]["IdApplicazione"];
                            btn.Size = new Size(200, 42);
                            btn.FlatStyle = FlatStyle.Flat;
                            btn.FlatAppearance.BorderSize = 0;
                            btn.Margin = new Padding(0, 0, 0, 0);
                            btn.Padding = new Padding(0, 0, 0, 0);
                            btn.ImageAlign = ContentAlignment.MiddleLeft;
                            btn.TextAlign = ContentAlignment.MiddleLeft;
                            btn.Dock = DockStyle.Top;
                            btn.Click += StartApplication;

                            menuLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, btn.Height));
                            menuLayout.Controls.Add(btn, 0, j++);

                            ToolStripItem item = menuIconTray.Items.Add(btn.Image);
                            item.Text = btn.Text;
                            item.Name = btn.Name;
                            item.Tag = btn.Tag;
                            item.Margin = new Padding(0, 0, 0, 0);
                            item.Padding = new Padding(0, 0, 0, 0);
                            item.Click += StartApplication;
                        }

                        var maxWidth = menuLayout.Controls
                            .OfType<Button>()
                            .Select(btn => btn.GetPreferredSize(btn.Size).Width)
                            .Max();

                        menuLayout.Width = maxWidth;
                    }
                }

                return true;
            }

            return false;
        }

        private object GetTag(object sender) 
        {
            Button button = sender as Button;
            ToolStripItem tsi = sender as ToolStripItem;

            if (button != null)
            return button.Tag;
            if (tsi != null)
            return tsi.Tag;

            throw new ArgumentException("Unexpected sender");
        }

        private void StartApplication(object sender, EventArgs e)
        {
            object tag = GetTag(sender);

            string path = @"%USERPROFILE%\PSO\";

            switch ((int)tag)
            {
                case 1:
                    path += "OfferteMGP";
                    break;
                case 2:
                case 3:
                case 4:
                case 13:
                    path += "InvioProgrammi";
                    break;
                case 5:
                    path += "ProgrammazioneImpianti";
                    break;
                case 6:
                    path += "UnitCommitment";
                    break;
                case 7:
                    path += "PrezziMSD";
                    break;
                case 8:
                    path += "SistemaComandi";
                    break;
                case 9:
                    path += "OfferteMSD";
                    break;
                case 10:
                    path += "OfferteMB";
                    break;
                case 11:
                    path += "ValidazioneTL";
                    break;
                case 12:
                    path += "PrevisioneCT";
                    break;
            }
            path = Environment.ExpandEnvironmentVariables(path);
#if DEBUG
            MessageBox.Show(path);
#endif
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            this.Opacity = 0d;
            this.ShowInTaskbar = false;
#if !DEBUG
            e.Cancel = true;
#endif
        }

        private void IconTray_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                CenterToScreen();
                menuLayout.Focus();
                this.ShowInTaskbar = true;
                this.Opacity = 100d;
            }
        }

    }
}
