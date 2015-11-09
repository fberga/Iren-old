using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Utility;
//using IWshRuntimeLibrary;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Deployment.Application;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.ToolsExcel.PSOLauncher
{
    public partial class Launcher : Form
    {
        #region Variabili
        
        private Excel.Application _xlLauncherApp = null;
        private ImageList _listaImmaginiApplicazioni = new ImageList();
        private bool _error = false;
        private string _errorMsg = "";
        
        #endregion

        #region Proprietà
        
        private string ErrorMsg 
        {
            get { return _errorMsg; }
            set
            {
                _errorMsg = value;
                MessageBox.Show(_errorMsg, "PSO - ERRORE!!!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public int IdUtente 
        { 
            get; 
            private set; 
        }
        
        #endregion

        #region Costruttore
        
        public Launcher() 
        {
            InitializeComponent();

#if !DEBUG
            //copio file nella cartella di esecuzione automatica
            string path = Environment.ExpandEnvironmentVariables(@"%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup");
            var app = System.Reflection.Assembly.GetExecutingAssembly();

            try
            {
                Text += " - v." + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4);
            }
            catch { }

            var wsh = new IWshRuntimeLibrary.IWshShell_Class();
            IWshRuntimeLibrary.IWshShortcut shortcut = wsh.CreateShortcut(Path.Combine(path, "PSOLauncher.lnk")) as IWshRuntimeLibrary.IWshShortcut;
            shortcut.TargetPath = app.CodeBase;
            shortcut.Save();
#endif
            var resourceSet = Iren.ToolsExcel.Base.Properties.Resources.ResourceManager.GetResourceSet(CultureInfo.InstalledUICulture, true, true);
            var imgs =
                from r in resourceSet.Cast<DictionaryEntry>()
                where r.Value is Image
                select r;

            foreach (var img in imgs)
                _listaImmaginiApplicazioni.Images.Add(img.Key as string, img.Value as Image);

            _listaImmaginiApplicazioni.ImageSize = new System.Drawing.Size(32, 32);

            _xlLauncherApp = new Excel.Application();
            _xlLauncherApp.Caption = "PSO";
        }
        
        #endregion

        #region Eventi
        
        private void Launcher_Load(object sender, EventArgs e) 
        {
            DataBase.CreateNew(ConfigurationManager.AppSettings["DB"]);

            if (!InitUsr())
            {
                _error = true;
                ErrorMsg = "L'utente non è configurato per l'utilizzo della suite PSO. Contattare l'amministratore.";
                return;
            }

            DataBase.IdUtente = IdUtente;

            if (!InitApplicazioni())
            {
                _error = true;
                ErrorMsg = "In seguito ad un errore, non è stato possibile caricare la lista delle applicazioni configurate. Contattare l'amministratore.";
                return;
            }
            _error = false;
        }
        private void StartApplication(object sender, EventArgs e) 
        {
            if (_xlLauncherApp == null)
            {
                _xlLauncherApp = new Excel.Application();
                _xlLauncherApp.Caption = "PSO";
            }

            int idApplicazione = (int)GetTag(sender);
            Workbook.AvviaApplicazione(_xlLauncherApp, idApplicazione);
        }
        private void IconTray_MouseDoubleClick(object sender, MouseEventArgs e) 
        {
            if (_error)
                ErrorMsg = _errorMsg;
            else if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                CenterToScreen();
                BringToFront();
                menuLayout.Focus();
                this.ShowInTaskbar = true;
                this.Opacity = 100d;
            }
        }
        private void IconTray_MouseClick(object sender, MouseEventArgs e) 
        {
            if (_error)
            {
                if (e.Button == System.Windows.Forms.MouseButtons.Right)
                {
                    if (!menuIconTray.Items.ContainsKey("Ricarica"))
                    {
                        menuIconTray.Items.Clear();

                        ToolStripItem ricarica = menuIconTray.Items.Add("Ricarica launcher");
                        ricarica.Name = "Ricarica";
                        ricarica.Margin = new Padding(0, 0, 0, 0);
                        ricarica.Padding = new Padding(0, 0, 0, 0);
                        ricarica.Click += RicaricaApplicazione;
                    }
                }
            }
        }
        private void RicaricaApplicazione(object sender, EventArgs e) 
        {
            if (MessageBox.Show("Ricaricare l'applicazione?", "PSO - ATTENZIONE!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                this.Launcher_Load(sender, e);
        }
        private void Launcher_FormClosing(object sender, FormClosingEventArgs e) 
        {
            this.Opacity = 0d;
            this.ShowInTaskbar = false;
#if !DEBUG
            if(e.CloseReason == CloseReason.UserClosing)
                e.Cancel = true;
            else
            {
                try
                {
                    IconTray.Icon = null;
                    //foreach (Excel.Workbook wb in _xlLauncherApp.Workbooks)
                    //    wb.Close();
                    _xlLauncherApp.Quit();
                }
                catch { }
            }
#else
            _xlLauncherApp.Quit();
#endif
        } 
        
        #endregion

        #region Metodi

        private bool InitUsr() 
        {
            DataTable dtUtente = DataBase.Select(DataBase.SP.UTENTE, "@CodUtenteWindows=" + Environment.UserName);

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
                menuLayout.Controls.Clear();
                menuIconTray.Items.Clear();

                int j = 0;
                for (int i = 0; i < dtApplicazioni.Rows.Count; i++)
                {
                    DataTable dtControllo = DataBase.Select(DataBase.SP.RIBBON.CONTROLLO_APPLICAZIONE, "@IdApplicazione=" + dtApplicazioni.Rows[i]["IdApplicazione"]);

                    if (dtControllo != null && dtControllo.Rows.Count > 0)
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
        private bool IsWorkbookOpen(string name) 
        {
            try
            {
                Microsoft.Office.Interop.Excel._Workbook wb = _xlLauncherApp.Workbooks[name + ".xlsm"];
                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion
    }
}
