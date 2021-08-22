namespace Stimulsoft.Report.Dictionary
{
    partial class OracleConnectionEditor
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
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.labelServerName = new System.Windows.Forms.Label();
            this.cbOraServer = new System.Windows.Forms.ComboBox();
            this.propertyGrid_Advanced = new System.Windows.Forms.PropertyGrid();
            this.tOraPassword = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tOraUser = new System.Windows.Forms.TextBox();
            this.labelName = new System.Windows.Forms.Label();
            this.bTest = new System.Windows.Forms.Button();
            this.bOk = new System.Windows.Forms.Button();
            this.bCancel = new System.Windows.Forms.Button();
            this.groupBoxLogOnToTheServer = new System.Windows.Forms.GroupBox();
            this.groupBoxAdvanced = new Stimulsoft.Controls.StiGroupBox();
            this.groupBoxLogOnToTheServer.SuspendLayout();
            this.groupBoxAdvanced.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelServerName
            // 
            this.labelServerName.AutoSize = true;
            this.labelServerName.Location = new System.Drawing.Point(12, 9);
            this.labelServerName.Name = "labelServerName";
            this.labelServerName.Size = new System.Drawing.Size(72, 13);
            this.labelServerName.TabIndex = 0;
            this.labelServerName.Text = "Server Name:";
            // 
            // cbOraServer
            // 
            this.cbOraServer.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append;
            this.cbOraServer.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbOraServer.FormattingEnabled = true;
            this.cbOraServer.Location = new System.Drawing.Point(12, 25);
            this.cbOraServer.Name = "cbOraServer";
            this.cbOraServer.Size = new System.Drawing.Size(251, 21);
            this.cbOraServer.Sorted = true;
            this.cbOraServer.TabIndex = 1;
            // 
            // propertyGrid_Advanced
            // 
            this.propertyGrid_Advanced.Location = new System.Drawing.Point(6, 19);
            this.propertyGrid_Advanced.Name = "propertyGrid_Advanced";
            this.propertyGrid_Advanced.Size = new System.Drawing.Size(338, 403);
            this.propertyGrid_Advanced.TabIndex = 0;
            this.propertyGrid_Advanced.Enter += new System.EventHandler(this.propertyGrid_Advanced_Enter);
            this.propertyGrid_Advanced.PropertyValueChanged += new System.Windows.Forms.PropertyValueChangedEventHandler(this.propertyGrid_Advanced_PropertyValueChanged);
            // 
            // tOraPassword
            // 
            this.tOraPassword.Location = new System.Drawing.Point(165, 45);
            this.tOraPassword.Name = "tOraPassword";
            this.tOraPassword.PasswordChar = '*';
            this.tOraPassword.Size = new System.Drawing.Size(176, 20);
            this.tOraPassword.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 48);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Password:";
            // 
            // tOraUser
            // 
            this.tOraUser.Location = new System.Drawing.Point(165, 19);
            this.tOraUser.Name = "tOraUser";
            this.tOraUser.Size = new System.Drawing.Size(176, 20);
            this.tOraUser.TabIndex = 1;
            // 
            // labelName
            // 
            this.labelName.AutoSize = true;
            this.labelName.Location = new System.Drawing.Point(8, 22);
            this.labelName.Name = "labelName";
            this.labelName.Size = new System.Drawing.Size(38, 13);
            this.labelName.TabIndex = 0;
            this.labelName.Text = "Name:";
            // 
            // bTest
            // 
            this.bTest.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.bTest.Location = new System.Drawing.Point(12, 158);
            this.bTest.Name = "bTest";
            this.bTest.Size = new System.Drawing.Size(114, 23);
            this.bTest.TabIndex = 4;
            this.bTest.Text = "Test Connection";
            this.bTest.UseVisualStyleBackColor = true;
            this.bTest.Click += new System.EventHandler(this.bTest_Click);
            // 
            // bOk
            // 
            this.bOk.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bOk.Location = new System.Drawing.Point(186, 158);
            this.bOk.Name = "bOk";
            this.bOk.Size = new System.Drawing.Size(75, 23);
            this.bOk.TabIndex = 5;
            this.bOk.Text = "Ok";
            this.bOk.UseVisualStyleBackColor = true;
            this.bOk.Click += new System.EventHandler(this.bOk_Click);
            // 
            // bCancel
            // 
            this.bCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.bCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.bCancel.Location = new System.Drawing.Point(267, 158);
            this.bCancel.Name = "bCancel";
            this.bCancel.Size = new System.Drawing.Size(75, 23);
            this.bCancel.TabIndex = 6;
            this.bCancel.Text = "Cancel";
            this.bCancel.UseVisualStyleBackColor = true;
            // 
            // groupBoxLogOnToTheServer
            // 
            this.groupBoxLogOnToTheServer.Controls.Add(this.tOraUser);
            this.groupBoxLogOnToTheServer.Controls.Add(this.tOraPassword);
            this.groupBoxLogOnToTheServer.Controls.Add(this.labelName);
            this.groupBoxLogOnToTheServer.Controls.Add(this.label2);
            this.groupBoxLogOnToTheServer.Location = new System.Drawing.Point(4, 52);
            this.groupBoxLogOnToTheServer.Name = "groupBoxLogOnToTheServer";
            this.groupBoxLogOnToTheServer.Size = new System.Drawing.Size(350, 73);
            this.groupBoxLogOnToTheServer.TabIndex = 2;
            this.groupBoxLogOnToTheServer.TabStop = false;
            this.groupBoxLogOnToTheServer.Text = "Log on to the server";
            // 
            // groupBoxAdvanced
            // 
            this.groupBoxAdvanced.AllowCollapse = true;
            this.groupBoxAdvanced.Collapsed = true;
            this.groupBoxAdvanced.Controls.Add(this.propertyGrid_Advanced);
            this.groupBoxAdvanced.Location = new System.Drawing.Point(4, 131);
            this.groupBoxAdvanced.Name = "groupBoxAdvanced";
            this.groupBoxAdvanced.ResHeight = 428;
            this.groupBoxAdvanced.Size = new System.Drawing.Size(350, 20);
            this.groupBoxAdvanced.TabIndex = 3;
            this.groupBoxAdvanced.TabStop = false;
            this.groupBoxAdvanced.Text = "Advanced";
            // 
            // OracleConnectionEditor
            // 
            this.AcceptButton = this.bOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.bCancel;
            this.ClientSize = new System.Drawing.Size(358, 188);
            this.Controls.Add(this.groupBoxAdvanced);
            this.Controls.Add(this.labelServerName);
            this.Controls.Add(this.groupBoxLogOnToTheServer);
            this.Controls.Add(this.cbOraServer);
            this.Controls.Add(this.bCancel);
            this.Controls.Add(this.bOk);
            this.Controls.Add(this.bTest);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "OracleConnectionEditor";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Oracle ODP.NET Connection";
            this.groupBoxLogOnToTheServer.ResumeLayout(false);
            this.groupBoxLogOnToTheServer.PerformLayout();
            this.groupBoxAdvanced.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelServerName;
        private System.Windows.Forms.TextBox tOraPassword;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tOraUser;
        private System.Windows.Forms.Label labelName;
        private System.Windows.Forms.ComboBox cbOraServer;
        private System.Windows.Forms.Button bTest;
        private System.Windows.Forms.PropertyGrid propertyGrid_Advanced;
        private System.Windows.Forms.Button bOk;
        private System.Windows.Forms.Button bCancel;
        private System.Windows.Forms.GroupBox groupBoxLogOnToTheServer;
        private Stimulsoft.Controls.StiGroupBox groupBoxAdvanced;
    }
}