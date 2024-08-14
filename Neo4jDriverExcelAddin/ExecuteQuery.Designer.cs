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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.LoadButton = new System.Windows.Forms.Button();
            this.UpdateButton = new System.Windows.Forms.Button();
            this.SyncAllButton = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
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
            this.txtCypher.Size = new System.Drawing.Size(629, 176);
            this.txtCypher.TabIndex = 0;
            // 
            // btnExecute
            // 
            this.btnExecute.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExecute.Location = new System.Drawing.Point(359, 20);
            this.btnExecute.Margin = new System.Windows.Forms.Padding(2);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(109, 47);
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
            this.txtNeoResponse.Location = new System.Drawing.Point(0, 364);
            this.txtNeoResponse.Multiline = true;
            this.txtNeoResponse.Name = "txtNeoResponse";
            this.txtNeoResponse.ReadOnly = true;
            this.txtNeoResponse.Size = new System.Drawing.Size(634, 135);
            this.txtNeoResponse.TabIndex = 2;
            this.txtNeoResponse.TextChanged += new System.EventHandler(this.txtNeoResponse_TextChanged);
            // 
            // connectionaddress
            // 
            this.connectionaddress.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.connectionaddress.Location = new System.Drawing.Point(3, 4);
            this.connectionaddress.Name = "connectionaddress";
            this.connectionaddress.Size = new System.Drawing.Size(487, 20);
            this.connectionaddress.TabIndex = 3;
            this.connectionaddress.Text = "bolt://localhost:7687/";
            // 
            // connectButton
            // 
            this.connectButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.connectButton.Location = new System.Drawing.Point(512, 3);
            this.connectButton.Name = "connectButton";
            this.connectButton.Size = new System.Drawing.Size(119, 20);
            this.connectButton.TabIndex = 4;
            this.connectButton.Text = "Connect";
            this.connectButton.UseVisualStyleBackColor = true;
            this.connectButton.Click += new System.EventHandler(this.connectButton_Click);
            // 
            // CreateNodeButton
            // 
            this.CreateNodeButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.CreateNodeButton.Location = new System.Drawing.Point(0, 19);
            this.CreateNodeButton.Name = "CreateNodeButton";
            this.CreateNodeButton.Size = new System.Drawing.Size(84, 47);
            this.CreateNodeButton.TabIndex = 5;
            this.CreateNodeButton.Text = "Create Nodes";
            this.CreateNodeButton.UseVisualStyleBackColor = true;
            this.CreateNodeButton.Click += new System.EventHandler(this.CreateNodeButton_Click);
            // 
            // ExecuteCypherRowsButton
            // 
            this.ExecuteCypherRowsButton.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ExecuteCypherRowsButton.Location = new System.Drawing.Point(216, 19);
            this.ExecuteCypherRowsButton.Name = "ExecuteCypherRowsButton";
            this.ExecuteCypherRowsButton.Size = new System.Drawing.Size(138, 48);
            this.ExecuteCypherRowsButton.TabIndex = 7;
            this.ExecuteCypherRowsButton.Text = "Execute Selection";
            this.ExecuteCypherRowsButton.UseVisualStyleBackColor = true;
            this.ExecuteCypherRowsButton.Click += new System.EventHandler(this.ExecuteCypherRowsButton_Click);
            // 
            // CreateRelationshipsButton
            // 
            this.CreateRelationshipsButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.CreateRelationshipsButton.Location = new System.Drawing.Point(90, 19);
            this.CreateRelationshipsButton.Name = "CreateRelationshipsButton";
            this.CreateRelationshipsButton.Size = new System.Drawing.Size(120, 47);
            this.CreateRelationshipsButton.TabIndex = 6;
            this.CreateRelationshipsButton.Text = "Create Relationships";
            this.CreateRelationshipsButton.UseVisualStyleBackColor = true;
            this.CreateRelationshipsButton.Click += new System.EventHandler(this.CreateRelationshipsButton_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Enabled = false;
            this.progressBar1.Location = new System.Drawing.Point(496, 4);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(10, 20);
            this.progressBar1.TabIndex = 8;
            this.progressBar1.Click += new System.EventHandler(this.progressBar1_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.SyncAllButton);
            this.groupBox1.Controls.Add(this.LoadButton);
            this.groupBox1.Controls.Add(this.UpdateButton);
            this.groupBox1.Controls.Add(this.CreateNodeButton);
            this.groupBox1.Controls.Add(this.CreateRelationshipsButton);
            this.groupBox1.Controls.Add(this.ExecuteCypherRowsButton);
            this.groupBox1.Controls.Add(this.btnExecute);
            this.groupBox1.Location = new System.Drawing.Point(0, 210);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(634, 148);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Controls";
            // 
            // LoadButton
            // 
            this.LoadButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.LoadButton.Location = new System.Drawing.Point(143, 72);
            this.LoadButton.Margin = new System.Windows.Forms.Padding(2);
            this.LoadButton.Name = "LoadButton";
            this.LoadButton.Size = new System.Drawing.Size(109, 47);
            this.LoadButton.TabIndex = 10;
            this.LoadButton.Text = "Pull";
            this.LoadButton.UseVisualStyleBackColor = true;
            this.LoadButton.Click += new System.EventHandler(this.LoadButton_Click);
            // 
            // UpdateButton
            // 
            this.UpdateButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.UpdateButton.Location = new System.Drawing.Point(2, 71);
            this.UpdateButton.Margin = new System.Windows.Forms.Padding(2);
            this.UpdateButton.Name = "UpdateButton";
            this.UpdateButton.Size = new System.Drawing.Size(109, 47);
            this.UpdateButton.TabIndex = 9;
            this.UpdateButton.Text = "Push";
            this.UpdateButton.UseVisualStyleBackColor = true;
            this.UpdateButton.Click += new System.EventHandler(this.UpdateButton_Click);
            // 
            // SyncAllButton
            // 
            this.SyncAllButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.SyncAllButton.Location = new System.Drawing.Point(359, 72);
            this.SyncAllButton.Margin = new System.Windows.Forms.Padding(2);
            this.SyncAllButton.Name = "SyncAllButton";
            this.SyncAllButton.Size = new System.Drawing.Size(109, 47);
            this.SyncAllButton.TabIndex = 11;
            this.SyncAllButton.Text = "Sync All";
            this.SyncAllButton.UseVisualStyleBackColor = true;
            this.SyncAllButton.Click += new System.EventHandler(this.SyncAllButton_Click);
            // 
            // ExecuteQuery
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.connectButton);
            this.Controls.Add(this.connectionaddress);
            this.Controls.Add(this.txtNeoResponse);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.txtCypher);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "ExecuteQuery";
            this.Size = new System.Drawing.Size(634, 498);
            this.groupBox1.ResumeLayout(false);
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
        internal EventHandler<SelectionArgs> LoadButtonEventHandler;
        internal EventHandler SyncAllButtonEventHandler;
        internal EventHandler<SelectionArgs> UpdateButtonEventHandler;

        private System.Windows.Forms.TextBox txtNeoResponse;
        private System.Windows.Forms.TextBox connectionaddress;
        private System.Windows.Forms.Button connectButton;
        private System.Windows.Forms.Button CreateNodeButton;
        private System.Windows.Forms.ToolTip createNodeTooltip;
        private System.Windows.Forms.Button ExecuteCypherRowsButton;
        private System.Windows.Forms.Button CreateRelationshipsButton;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button UpdateButton;
        private System.Windows.Forms.Button LoadButton;
        private System.Windows.Forms.Button SyncAllButton;
    }
}
