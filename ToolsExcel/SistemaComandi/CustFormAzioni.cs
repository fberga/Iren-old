using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel
{
    class CustFormAzioni : Forms.FormAzioni
    {

        public override bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime? dataRif = null)
        {
            System.Windows.Forms.MessageBox.Show("ciao ciao");

            return base.EsportaAzioneInformazione(siglaEntita, siglaAzione, desEntita, desAzione, dataRif);
        }

    }
}
