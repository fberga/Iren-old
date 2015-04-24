using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Iren.ToolsExcel.Forms
{
    public partial class LoaderScreen : Form
    {
        BackgroundWorker bkWorker = new BackgroundWorker();

        public string LoadingWhat
        {
            set
            {
                lbText.Text = value;
                this.Refresh();
            }
        }


        public LoaderScreen()
        {
            InitializeComponent();
            bkWorker.WorkerReportsProgress = true;
            bkWorker.WorkerSupportsCancellation = true;

            bkWorker.DoWork += bkWorker_DoWork;
            bkWorker.ProgressChanged += bkWorker_ProgressChanged;
        }

        private void lbText_SizeChanged(object sender, EventArgs e)
        {
            double width = lbText.Width;
            //double height = lbText.Height;

            //lbText.Top = (int)Math.Round((this.Height / 2) - (height / 2));
            lbText.Left = (int)Math.Round((this.Width / 2) - (width / 2));
        }

        private void LoaderScreen_Load(object sender, EventArgs e)
        {
        }

        void bkWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            string s = "";
            
            for (int j = 0; j < e.ProgressPercentage; j++)
                s += ".";

            lbCaricamento.Invoke((MethodInvoker)delegate
            {
                lbCaricamento.Text = "Caricamento " + s;
            });
            this.Invoke((MethodInvoker)delegate
            {
                this.Refresh();
            });
        }

        void bkWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            int i = 0;
            while (true)
            {
                if ((worker.CancellationPending == true))
                {
                    e.Cancel = true;
                    break;
                }
                else
                {
                    i = (i % 3) + 1;

                    worker.ReportProgress(i);
                    System.Threading.Thread.Sleep(500);
                }
            }
        }

        private void LoaderScreen_VisibleChanged(object sender, EventArgs e)
        {
            if (this.Visible)
            {
                bkWorker.RunWorkerAsync();
            }
            else
            {
                bkWorker.CancelAsync();
            }
        }
    }
}
