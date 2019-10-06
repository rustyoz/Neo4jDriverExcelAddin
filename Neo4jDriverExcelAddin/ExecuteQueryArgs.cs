namespace Neo4jDriverExcelAddin
{
    using System;

    internal class ExecuteQueryArgs : EventArgs
    {
        public string Cypher { get; set; }
    }

    
}