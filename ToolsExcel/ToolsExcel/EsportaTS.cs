using Iren.ToolsExcel.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel
{
    class EsportaTS : IEsporta
    {
        public bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime? dataRif = null)
        {
            System.Windows.Forms.MessageBox.Show("ciao");

            return true;
        }
    }
}
