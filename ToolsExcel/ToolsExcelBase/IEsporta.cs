using System;

namespace Iren.ToolsExcel.Base
{
    public interface IEsporta
    {
        bool EsportaAzioneInformazione(object siglaEntita, object siglaAzione, object desEntita, object desAzione, DateTime? dataRif = null);
    }
}
