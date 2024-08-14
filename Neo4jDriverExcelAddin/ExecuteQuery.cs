﻿namespace Neo4jDriverExcelAddin
{
    using System;
    using System.Windows.Forms;

    public partial class ExecuteQuery : UserControl
    {
        private delegate void SafeCallDelegate(string text);
        public IProgress<int> progress;
        public ExecuteQuery()
        {
            InitializeComponent();
            createNodeTooltip.SetToolTip(CreateNodeButton, "Create Nodes with header row as label and properties.");

            progress = new Progress<int>(v =>
            {
                progressBar1.Value = v;
            });
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

            ExecuteCypher?.Invoke(this, new ExecuteQueryArgs { Cypher = cypher });
        }

        public void SetMessage(string message)
        {
            if(txtNeoResponse.InvokeRequired)
            {
                var d = new SafeCallDelegate(SetMessage);
                txtNeoResponse.Invoke(d,  new object[] { message });

            }
            else
            {
                txtNeoResponse.Text = message;

            }
                
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

        private void CreateRelationshipsButton_Click(object sender, EventArgs e)
        {
            CreateRelationships?.Invoke(this, new SelectionArgs { SelectionRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection });
        }

        private void UpdateButton_Click(object sender, EventArgs e)
        {
            UpdateButtonEventHandler?.Invoke(this, new SelectionArgs { SelectionRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection });
        }

        private void LoadButton_Click(object sender, EventArgs e)
        {
            LoadButtonEventHandler?.Invoke(this, new SelectionArgs { SelectionRange = Globals.ThisAddIn.Application.ActiveWindow.RangeSelection });
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void txtNeoResponse_TextChanged(object sender, EventArgs e)
        {

        }

        private void SyncAllButton_Click(object sender, EventArgs e)
        {
            SyncAllButtonEventHandler?.Invoke(this,EventArgs.Empty);
        }
    }
}