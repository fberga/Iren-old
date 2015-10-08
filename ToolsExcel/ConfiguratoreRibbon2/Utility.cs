using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    class SP
    {
        public const string CONTROLLO = "RIBBON.spControllo";
        public const string APPLICAZIONI = "spApplicazioneProprieta";
        public const string INSERT_CONTROLLO = "RIBBON.spInsertControllo";
        public const string INSERT_GRUPPO = "RIBBON.spInsertGruppo";
        public const string INSERT_GRUPPO_CONTROLLO = "RIBBON.spInsertGruppoControllo";
        public const string APPLICAZIONE_UTENTE_RIBBON = "RIBBON.spGruppoControllo";
    }

    class Utility
    {
        public static IEnumerable<Control> GetAll(Control control, Type type = null)
        {
            var controls = control.Controls.Cast<Control>();

            return controls.SelectMany(ctrl => GetAll(ctrl, type))
                                      .Concat(controls)
                                      .Where(c => type == null || c.GetType() == type || c.GetType().GetInterfaces().Contains(type));
        }

        public static int FindLastOfItsKind(Control ctrl, string prefix, Type type)
        {
            var progs = GetAll(ctrl, type)
                .Where(c => c.Name.StartsWith(PrepareLabelForControlName(prefix)))
                .Select(c =>
                {
                    string num = Regex.Match(c.Name, @"\d+").Value;
                    int progNum = 0;
                    int.TryParse(num, out progNum);
                    return progNum;
                }).ToList();

            if (progs.Count > 0)
                return progs.Max();

            return 0;
        }

        public static SizeF MeasureTextSize(Control ctrl)
        {
            //calcolo la dimensione
            //lavoro su 2 righe...quindi calcolo tutte le dimensioni delle parole e poi le combino per tentativi mettendo:
            // 1 sopra, tot - 1 sotto; 2 sopra, tot - 2 sotto; ... 

            string s = ctrl.Text;

            //se è un tasto a dimensione piccola, calcolo normalmente
            int dim = 1;
            if (ctrl.GetType() == typeof(RibbonButton))
                dim = ((RibbonButton)ctrl).Dimensione;

            if (!s.Contains(' ') || ctrl.GetType() == typeof(TextBox) || dim == 0)
                return ctrl.CreateGraphics().MeasureString(s, ctrl.Font, int.MaxValue);

            string[] parole = s.Split(' ');
            float[] misure = new float[parole.Length];

            //calcolo le singole dimensioni
            for (int i = 0; i < parole.Length; i++)
                misure[i] = ctrl.CreateGraphics().MeasureString(parole[i], ctrl.Font, int.MaxValue).Width;

            //provo a combinare tutte le parole e vedo quale combinazione mi da dimensione minima (forse anche rapporto più bilanciato...)
            float riga1 = Enumerable.Sum(misure);
            float riga2 = 0;

            //float rapporto = 0;
            float opt = riga1;

            //ciclo ma lascio almeno una parole sopra
            for (int i = parole.Length - 1; i > 0; i--)
            {
                riga2 += misure[i];
                riga1 -= misure[i];

                float tmpOpt = Math.Max(riga1, riga2);

                if (opt > tmpOpt)
                {
                    opt = tmpOpt;
                }
            }

            return ctrl.CreateGraphics().MeasureString(s, ctrl.Font, (int)Math.Ceiling(opt));

        }

        public static void UpdateGroupDimension(Control parent)
        {
            var txtWidth =
                (from txt in parent.Controls.OfType<TextBox>()
                 select txt.GetPreferredSize(txt.Size)).FirstOrDefault();
                 //select (int)(Utility.MeasureTextSize(txt).Width + 20)).FirstOrDefault();

            var totWidth =
                (from p in parent.Controls.OfType<ControlContainer>()
                 select p.Width).DefaultIfEmpty().Sum() + 20;

            var containers = parent.Controls.OfType<ControlContainer>()
                .OrderBy(c => c.Left)
                .DefaultIfEmpty()
                .ToArray();

            for (int i = 1; i < containers.Length; i++)
                containers[i].Left = containers[i - 1].Right;


            parent.Width = Math.Max(txtWidth.Width, totWidth);
            parent.Invalidate();            
        }

        public static void GroupsDisplacement(Control ribbon)
        {
            var groups = ribbon.Controls.OfType<RibbonGroup>()
                //.OrderBy(g => g.Left)
                .ToList();

            if (groups.Count > 0)
            {
                int left = ribbon.Padding.Left;
                foreach (RibbonGroup group in groups)
                {
                    group.Left = left;
                    left = group.Right;
                }
            }            
        }

        public static string PrepareLabelForControlName(string label)
        {
            return label.Replace(" ", "");
        }
        
        public static Image GetResurceImage(string name)
        {
            return Iren.ToolsExcel.Base.Properties.Resources.ResourceManager.GetObject(name) as Image;
        }

        public static Control CreateEmptyContainer(Control parent)
        {
            ControlContainer container = new ControlContainer();
            container.Size = new Size(50, parent.Height - 30);

            var left =
                (from p in parent.Controls.OfType<ControlContainer>()
                 select p.Right).DefaultIfEmpty().Max();

            container.Left = left == 0 ? parent.Padding.Left : left + 10;
            container.Top = parent.Padding.Top;
            return container;
        }
    }
}
