namespace Neo4jDriverExcelAddin
{
    using System;

    partial class ExecuteQuery
    {
        /// <summary> 
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            if (disposing)
            {
                ExecuteCypher = null;
            }

            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.txtCypher = new System.Windows.Forms.TextBox();
            this.btnExecute = new System.Windows.Forms.Button();
            this.txtNeoResponse = new System.Windows.Forms.TextBox();
            this.connectionaddress = new System.Windows.Forms.TextBox();
            this.connectButton = new System.Windows.Forms.Button();
            this.CreateNodeButton = new System.Windows.Forms.Button();
            this.createNodeTooltip = new System.Windows.Forms.ToolTip(this.components);
            this.ExecuteCypherRowsButton = new System.Windows.Forms.Button();
            this.CreateRelationshipsButton = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // txtCypher
            // 
            this.txtCypher.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtCypher.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtCypher.Location = new System.Drawing.Point(2, 29);
            this.txtCypher.Margin = new System.Windows.Forms.Padding(2);
            this.txtCypher.Multiline = true;
            this.txtCypher.Name = "txtCypher";
            this.txtCypher.Size = new System.Drawing.Size(389, 176);
            this.txtCypher.TabIndex = 0;
            // 
            // btnExecute
            // 
            this.btnExecute.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExecute.Location = new System.Drawing.Point(311, 241);
            this.btnExecute.Margin = new System.Windows.Forms.Padding(2);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(76, 47);
            this.btnExecute.TabIndex = 1;
            this.btnExecute.Text = "Execute Cypher";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // txtNeoResponse
            // 
            this.txtNeoResponse.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtNeoResponse.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtNeoResponse.Location = new System.Drawing.Point(0, 293);
            this.txtNeoResponse.Multiline = true;
            this.txtNeoResponse.Name = "txtNeoResponse";
            this.txtNeoResponse.ReadOnly = true;
            this.txtNeoResponse.Size = new System.Drawing.Size(389, 206);
            this.txtNeoResponse.TabIndex = 2;
            // 
            // connectionaddress
            // 
            this.connectionaddress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.connectionaddress.Location = new System.Drawing.Point(3, 4);
            this.connectionaddress.Name = "connectionaddress";
            this.connectionaddress.Size = new System.Drawing.Size(302, 20);
            this.connectionaddress.TabIndex = 3;
            this.connectionaddress.Text = "bolt://localhost:7687/";
            // 
            // connectButton
            // 
            this.connectButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.connectButton.Location = new System.Drawing.Point(311, 3);
            this.connectButton.Name = "connectButton";
            this.connectButton.Size = new System.Drawing.Size(75, 20);
            this.connectButton.TabIndex = 4;
            this.connectButton.Text = "Connect";
            this.connectButton.UseVisualStyleBackColor = true;
            this.connectButton.Click += new System.EventHandler(this.connectButton_Click);
            // 
            // CreateNodeButton
            // 
            this.CreateNodeButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.CreateNodeButton.Location = new System.Drawing.Point(3, 239);
            this.CreateNodeButton.Name = "CreateNodeButton";
            this.CreateNodeButton.Size = new System.Drawing.Size(57, 47);
            this.CreateNodeButton.TabIndex = 5;
            this.CreateNodeButton.Text = "Create Nodes";
            this.CreateNodeButton.UseVisualStyleBackColor = true;
            this.CreateNodeButton.Click += new System.EventHandler(this.CreateNodeButton_Click);
            // 
            // ExecuteCypherRowsButton
            // 
            this.ExecuteCypherRowsButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ExecuteCypherRowsButton.Location = new System.Drawing.Point(151, 239);
            this.ExecuteCypherRowsButton.Name = "ExecuteCypherRowsButton";
            this.ExecuteCypherRowsButton.Size = new System.Drawing.Size(94, 48);
            this.ExecuteCypherRowsButton.TabIndex = 7;
            this.ExecuteCypherRowsButton.Text = "Execute Selection";
            this.ExecuteCypherRowsButton.UseVisualStyleBackColor = true;
            this.ExecuteCypherRowsButton.Click += new System.EventHandler(this.ExecuteCypherRowsButton_Click);
            // 
            // CreateRelationshipsButton
            // 
            this.CreateRelationshipsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.CreateRelationshipsButton.Location = new System.Drawing.Point(66, 239);
            this.CreateRelationshipsButton.Name = "CreateRelationshipsButton";
            this.CreateRelationshipsButton.Size = new System.Drawing.Size(79, 47);
            this.CreateRelationshipsButton.TabIndex = 6;
            this.CreateRelationshipsButton.Text = "Create Relationships";
            this.CreateRelationshipsButton.UseVisualStyleBackColor = true;
            this.CreateRelationshipsButton.Click += new System.EventHandler(this.CreateRelationshipsButton_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(3, 210);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(383, 26);
            this.progressBar1.TabIndex = 8;
            // 
            // ExecuteQuery
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.ExecuteCypherRowsButton);
            this.Controls.Add(this.CreateNodeButton);
            this.Controls.Add(this.CreateRelationshipsButton);
            this.Controls.Add(this.connectButton);
            this.Controls.Add(this.connectionaddress);
            this.Controls.Add(this.txtNeoResponse);
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.txtCypher);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "ExecuteQuery";
            this.Size = new System.Drawing.Size(389, 498);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtCypher;
        private System.Windows.Forms.Button btnExecute;

        internal EventHandler<ConnectDatabaseArgs> ConnectDatabase;
        internal EventHandler<ExecuteQueryArgs> ExecuteCypher;
        internal EventHandler<SelectionArgs> CreateNodes;
        internal EventHandler<SelectionArgs> ExecuteSelection;
        internal EventHandler<SelectionArgs> CreateRelationships;

        private System.Windows.Forms.TextBox txtNeoResponse;
        private System.Windows.Forms.TextBox connectionaddress;
        private System.Windows.Forms.Button connectButton;
        private System.Windows.Forms.Button CreateNodeButton;
        private System.Windows.Forms.ToolTip createNodeTooltip;
        private System.Windows.Forms.Button ExecuteCypherRowsButton;
        private System.Windows.Forms.Button CreateRelationshipsButton;
        private System.Windows.Forms.ProgressBar progressBar1;
    }
}
