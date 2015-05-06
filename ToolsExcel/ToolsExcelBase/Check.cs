using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.Base
{
    public abstract class Check
    {
        public abstract void ExecuteCheck(NewDefinedNames newNomiDefiniti, string siglaEntita, int type);
    }
}
