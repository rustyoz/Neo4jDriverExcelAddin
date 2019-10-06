using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Neo4jDriverExcelAddin
{
    class SelectionArgs : EventArgs
    {
        public Microsoft.Office.Interop.Excel.Range SelectionRange { get; set; }
    }
}
