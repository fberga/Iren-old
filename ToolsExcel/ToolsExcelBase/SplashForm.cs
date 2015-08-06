using Iren.ToolsExcel.Utility;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Iren.ToolsExcel.Base
{
    public partial class SplashForm : Form
    {
        private const int CS_DROPSHADOW = 0x20000;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }
        }

        private delegate void ShowDelegate();
        private delegate void CloseDelegate();
        private delegate void UpdateStatusDelegate(string status);

        public SplashForm()
        {
            InitializeComponent();
        }

        public void ShowSplashScreen()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new ShowDelegate(ShowSplashScreen));
                return;
            }
            
            this.ShowDialog();                

            if (!this.IsDisposed)
                Application.Run(this);
        }

        public void CloseSplashScreen()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new CloseDelegate(CloseSplashScreen));
                return;
            }

            this.Close();
            this.Dispose();
        }

        public void UdpateStatusText(string status)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new UpdateStatusDelegate(UdpateStatusText), status);
                return;
            }

            this.BringToFront();
            if (status.Length > 70)
                status = status.Substring(0, 67) + " ...";
            if(status != lbText.Text)
                lbText.Text = status;
        }

        private void lbText_SizeChanged(object sender, EventArgs e)
        {
            double width = lbText.Width;
            lbText.Left = (int)Math.Round((panelAll.Width / 2) - (width / 2));
        }

        static SplashForm sf = null;

        public static void ShowSplash()
        {
            if (sf == null)
            {
                sf = new SplashForm();
                sf.ShowSplashScreen();
            }
        }
        public static void CloseSplash()
        {
            if (sf != null)
            {
                sf.CloseSplashScreen();
                sf = null;
            }
        }
        public static void UpdateStatus(string status)
        {
            if (sf != null)
            {
                sf.UdpateStatusText(status);
            }
        }
    }

    public static class SplashScreen
    {
        public static void Show()
        {
            Thread splashthread = new Thread(new ThreadStart(SplashForm.ShowSplash));
            splashthread.IsBackground = true;
           
            splashthread.Start();
        }

        public static void Close()
        {
            SplashForm.CloseSplash();
        }

        public static void UpdateStatus(string status)
        {
            SplashForm.UpdateStatus(status);
        }
        
    }
}
