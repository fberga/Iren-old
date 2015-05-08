using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.Base
{
    public abstract class ACarica
    {
        public abstract bool RunCarica(object siglaEntita, object siglaAzione, DateTime dataRif);
    }

    public class Carica : ACarica
    {
        public override bool RunCarica(object siglaEntita, object siglaAzione, DateTime dataRif)
        {
            return true;
        }
    }
}
