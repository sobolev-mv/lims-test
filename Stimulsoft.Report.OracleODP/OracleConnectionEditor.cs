using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Oracle.DataAccess.Client;


namespace Stimulsoft.Report.Dictionary
{
    public partial class OracleConnectionEditor : Form
    {
        #region Properties
        private OracleConnectionStringBuilder oracleBuilder;
        public OracleConnectionStringBuilder OracleBuilder
        {
            get 
            { 
                return oracleBuilder; 
            }
        }
        #endregion


        public OracleConnectionEditor(string connString)
        {
            InitializeComponent();

            if (string.IsNullOrEmpty(connString))
                oracleBuilder = new OracleConnectionStringBuilder();
            else
                oracleBuilder = new OracleConnectionStringBuilder(connString);

            ParseConnectionString();
            GetOraServers();

            propertyGrid_Advanced.SelectedObject = oracleBuilder;
        }

        #region Methods

        private void GetOraServers()
        {
            OracleDataSourceEnumerator oraDSEnum = new OracleDataSourceEnumerator();
            try
            {
                DataTable dataTable = oraDSEnum.GetDataSources();
                foreach (DataRow row in dataTable.Rows)
                    cbOraServer.Items.Add(row[0]);
            }
            catch
            {
            }
        }

        private string GetConnectionString()
        {
            oracleBuilder.UserID = tOraUser.Text;
            oracleBuilder.Password = tOraPassword.Text;
            oracleBuilder.DataSource = cbOraServer.Text;
            return oracleBuilder.ToString();
        }

        private void ParseConnectionString()
        {
            tOraUser.Text = oracleBuilder.UserID;
            tOraPassword.Text = oracleBuilder.Password;
            cbOraServer.Text = oracleBuilder.DataSource;
        }
        #endregion

        #region Handlers
        private void bTest_Click(object sender, EventArgs e)
        {
            try
            {
                using (OracleConnection oracleConnection = new OracleConnection(GetConnectionString()))
                {
                    oracleConnection.Open();
                    MessageBox.Show("Test OK!");
                    oracleConnection.Close();
                }
            }
            catch (Exception exception)
            {
                StiLogService.Write(this.GetType(), exception);
                Stimulsoft.Base.StiExceptionProvider.Show(exception);
            }
        }

        private void bOk_Click(object sender, EventArgs e)
        {
            GetConnectionString();
            DialogResult = DialogResult.OK;
        }

        private void propertyGrid_Advanced_PropertyValueChanged(object s, PropertyValueChangedEventArgs e)
        {
            ParseConnectionString();
        }

        private void propertyGrid_Advanced_Enter(object sender, EventArgs e)
        {
            GetConnectionString();
        }
        #endregion

    }
}
