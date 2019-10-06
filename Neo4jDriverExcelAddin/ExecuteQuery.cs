namespace Neo4jDriverExcelAddin
{
    using System;
    using System.Windows.Forms;

    public partial class ExecuteQuery : UserControl
    {
        public ExecuteQuery()
        {
            InitializeComponent();
            createNodeTooltip.SetToolTip(CreateNodeButton, "Create Nodes with header row as label and properties.");
        }

        public string ConnectionString()
        {
            return connectionaddress.Text;
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            RaiseExecuteCypherEvent(txtCypher.Text);
        }

        private void RaiseExecuteCypherEvent(string cypher)
        {
            if (string.IsNullOrWhiteSpace(cypher))
                return;

            ExecuteCypher?.Invoke(this, new ExecuteQueryArgs {Cypher = cypher});
        }

        public void SetMessage(string message)
        {
            txtNeoResponse.Text = message;
        }

        private void connectButton_Click(object sender, EventArgs e)
        {
            RaiseConenctDatabaseEvent(connectionaddress.Text);
        }

        private void RaiseConenctDatabaseEvent(string connectstring)
        {
            if (string.IsNullOrWhiteSpace(connectstring))
                return;

            ConnectDatabase?.Invoke(this, new ConnectDatabaseArgs { ConnectionString = connectstring });

        }

        private void CreateNodeButton_Click(object sender, EventArgs e)
        {
            CreateNodes?.Invoke(this, new SelectionArgs { SelectionRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection });

        }

        private void ExecuteCypherRowsButton_Click(object sender, EventArgs e)
        {
            ExecuteSelection?.Invoke(this, new SelectionArgs { SelectionRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection });
        }
    }
}