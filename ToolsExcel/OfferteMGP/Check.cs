using Iren.ToolsExcel.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel
{
    class Check : Base.Check
    {
        public override void ExecuteCheck(NewDefinedNames newNomiDefiniti, string siglaEntita, int type)
        {
            System.Windows.Forms.MessageBox.Show("Ciao");
        }
    }
}
