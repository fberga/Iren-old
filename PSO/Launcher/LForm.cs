using Iren.PSO.Base;
using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Deployment.Application;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Launcher
{
    public partial class LForm : Form
    {
        #region Costruttore
        
        public LForm(ContextMenuStrip menu)
        {
            InitializeComponent();
#if !DEBUG
            Text = "PSO - v." + ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4);
#endif
            int j = 0;
            foreach (ToolStripMenuItem item in menu.Items)
            {
                Button btn = new Button();
                btn.ImageList = menu.ImageList;
                btn.ImageKey = item.ImageKey;
                btn.Text = item.Text;
                btn.TextImageRelation = TextImageRelation.ImageBeforeText;
                btn.Name = item.Name;
                btn.Tag = item.Tag;
                btn.Size = new Size(200, 42);
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;
                btn.Margin = new Padding(0, 0, 0, 0);
                btn.Padding = new Padding(0, 0, 0, 0);
                btn.ImageAlign = ContentAlignment.MiddleLeft;
                btn.TextAlign = ContentAlignment.MiddleLeft;
                btn.Dock = DockStyle.Top;
                btn.Click += LDaemon.StartApplication;

                menuLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, btn.Height));
                menuLayout.Controls.Add(btn, 0, j++);
            }
        }
        
        #endregion

        #region Eventi
        
        private void Launcher_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            e.Cancel = true;
        }
        
        #endregion
    }
}
