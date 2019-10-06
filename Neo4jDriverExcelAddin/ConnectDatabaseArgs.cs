using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Neo4jDriverExcelAddin
{
    internal class ConnectDatabaseArgs : EventArgs
    {
        public string ConnectionString { get; set; }
    }

}
