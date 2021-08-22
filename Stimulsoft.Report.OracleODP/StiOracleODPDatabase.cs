#region Copyright (C) 2003-2010 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports       										}
{																	}
{	Copyright (C) 2003-2010 Stimulsoft     							}
{	ALL RIGHTS RESERVED												}
{																	}
{	The entire contents of this file is protected by U.S. and		}
{	International Copyright Laws. Unauthorized reproduction,		}
{	reverse-engineering, and distribution of all or any portion of	}
{	the code contained in this file is strictly prohibited and may	}
{	result in severe civil and criminal penalties and will be		}
{	prosecuted to the maximum extent possible under the law.		}
{																	}
{	RESTRICTIONS													}
{																	}
{	THIS SOURCE CODE AND ALL RESULTING INTERMEDIATE FILES			}
{	ARE CONFIDENTIAL AND PROPRIETARY								}
{	TRADE SECRETS OF Stimulsoft										}
{																	}
{	CONSULT THE END USER LICENSE AGREEMENT FOR INFORMATION ON		}
{	ADDITIONAL RESTRICTIONS.										}
{																	}
{*******************************************************************}
*/
#endregion Copyright (C) 2003-2010 Stimulsoft

using System;
using System.Data;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections;
using Stimulsoft.Base.Localization;
using Stimulsoft.Base.Serializing;
using Stimulsoft.Report.Dictionary.Design;
using Oracle.DataAccess.Client;

namespace Stimulsoft.Report.Dictionary
{

    [TypeConverter(typeof(StiSqlDatabaseConverter))]
    public class StiOracleODPDatabase : StiSqlDatabase
    {
        #region StiService override
        public override string ServiceName
        {
            get
            {
                return StiLocalization.Get("Database", "DatabaseOracleODP");
            }
        }
        #endregion

        #region DataAdapter override
        protected override string DataAdapterType
        {
            get
            {
                return "Stimulsoft.Report.Dictionary.StiOracleODPAdapterService";
            }
        }
        #endregion

        /// <summary>
        /// Adds tables, views and stored procedures to report dictionary from database information.
        /// </summary>
        public override void ApplyDatabaseInformation(StiDatabaseInformation information, StiReport report)
        {
            #region Tables
            foreach (DataTable dataTable in information.Tables)
            {
                StiOracleODPSource source = new StiOracleODPSource(this.Name,
                    StiNameCreation.CreateName(report, dataTable.TableName, false, false, true));
                string table = dataTable.TableName;
                if (table.Trim().Contains(" ")) table = string.Format("[{0}]", table);
                source.SqlCommand = "select * from " + table;

                foreach (DataColumn dataColumn in dataTable.Columns)
                {
                    StiDataColumn column = new StiDataColumn(dataColumn.ColumnName, dataColumn.DataType);
                    source.Columns.Add(column);
                }
                report.Dictionary.DataSources.Add(source);
            }
            #endregion

            #region Views
            foreach (DataTable dataTable in information.Views)
            {
                StiOracleODPSource source = new StiOracleODPSource(this.Name,
                    StiNameCreation.CreateName(report, dataTable.TableName, false, false, true));
                string table = dataTable.TableName;
                if (table.Trim().Contains(" ")) table = string.Format("[{0}]", table);
                source.SqlCommand = "select * from " + table;

                foreach (DataColumn dataColumn in dataTable.Columns)
                {
                    StiDataColumn column = new StiDataColumn(dataColumn.ColumnName, dataColumn.DataType);
                    source.Columns.Add(column);
                }
                report.Dictionary.DataSources.Add(source);
            }
            #endregion

            #region StoredProcedures
            foreach (DataTable dataTable in information.StoredProcedures)
            {
                StiOracleODPSource source = new StiOracleODPSource(this.Name,
                    StiNameCreation.CreateName(report, dataTable.TableName, false, false, true));
                source.SqlCommand = "execute " + dataTable.TableName;

                foreach (DataColumn dataColumn in dataTable.Columns)
                {
                    StiDataParameter parameter = new StiDataParameter();
                    parameter.Name = dataColumn.ColumnName;
                    source.Parameters.Add(parameter);
                }
                report.Dictionary.DataSources.Add(source);
            }
            #endregion
        }

        /// <summary>
        /// Returns full database information.
        /// </summary>
        public override StiDatabaseInformation GetDatabaseInformation()
        {
            StiDatabaseInformation information = new StiDatabaseInformation();
            try
            {
                using (OracleConnection connection = new OracleConnection(this.ConnectionString))
                {
                    connection.Open();

                    #region Tables
                    DataTable tables = connection.GetSchema("Tables");

                    Hashtable tableHash = new Hashtable();
                    try
                    {

                        foreach (DataRow row in tables.Rows)
                        {
                            if ((row["TYPE"] != DBNull.Value && ((string)row["TYPE"]) == "System") ||
                                IsSystemOwner(row["OWNER"] as string)) continue;

                            DataTable table = new DataTable(row["OWNER"] as string + "." + row["TABLE_NAME"] as string);

                            tableHash[table.TableName] = table;
                            information.Tables.Add(table);
                        }
                    }
                    catch
                    {
                    }
                    #endregion

                    #region Views
                    DataTable views = connection.GetSchema("Views");
                    Hashtable viewHash = new Hashtable();
                    try
                    {
                        foreach (DataRow row in views.Rows)
                        {
                            if (IsSystemOwner(row["OWNER"] as string))
                                continue;

                            DataTable table = new DataTable(row["OWNER"] as string + "." + row["VIEW_NAME"] as string);

                            viewHash[table.TableName] = table;
                            information.Views.Add(table);
                        }
                    }
                    catch
                    {
                    }
                    #endregion

                    #region Columns
                    try
                    {
                        DataTable columns = connection.GetSchema("Columns");
                        foreach (DataRow row in columns.Rows)
                        {
                            if (IsSystemOwner(row["OWNER"] as string))
                                continue;

                            string columnName = row["COLUMN_NAME"] as string;
                            string tableName = row["OWNER"] as string + "." + row["TABLE_NAME"] as string;

                            if (tableHash[tableName] != null)
                            {
                                Type columnType = ConvertDbTypeToTypeInternal(row["DATATYPE"] as string);

                                DataColumn column = new DataColumn(columnName, columnType);
                                DataTable table = tableHash[tableName] as DataTable;
                                if (table != null)
                                {
                                    table.Columns.Add(column);
                                }
                            }
                            else if ((viewHash[tableName] != null))
                            {
                                Type columnType = ConvertDbTypeToTypeInternal(row["DATATYPE"] as string);

                                DataColumn column = new DataColumn(columnName, columnType);
                                DataTable table = viewHash[tableName] as DataTable;
                                if (table != null)
                                {
                                    table.Columns.Add(column);
                                }
                            }
                        }
                    }
                    catch
                    {
                    }
                    #endregion

                    #region Procedures
                    DataTable procedures = connection.GetSchema("Procedures");

                    Hashtable procedureHash = new Hashtable();
                    try
                    {

                        foreach (DataRow row in procedures.Rows)
                        {
                            if ((row["OWNER"] != DBNull.Value && ((string)row["OWNER"]) == "SYS") || IsSystemOwner(row["OWNER"] as string))
                                continue;

                            DataTable table = new DataTable(row["OWNER"] as string + "." + row["OBJECT_NAME"] as string);

                            procedureHash[table.TableName] = table;
                            information.StoredProcedures.Add(table);
                        }
                    }
                    catch
                    {
                    }
                    #endregion

                    connection.Close();
                }
                return information;
            }
            catch
            {
                return null;
            }
        }

        private Type ConvertDbTypeToTypeInternal(string dbType)
        {
            dbType = dbType.Replace(" ", "");

            if (dbType == "DATE" || dbType.Contains("TIMESTAMP"))
                return typeof(DateTime);
            if (dbType.Contains("INTERVALDAY"))
                return typeof(TimeSpan);
            if (dbType.Contains("INTERVALYEAR"))
                return typeof(long);

            switch (dbType)
            {
                case "BFILE":
                case "BLOB":
                case "LONGRAW":
                case "RAW":
                    return typeof(byte[]);

                case "FLOAT":
                case "NUMBER":
                    return typeof(decimal);

                default:
                    return typeof(string);

            }
        }
        private bool IsSystemOwner(string owner)
        {
            //('SYS','SYSMAN','SYSTEM','WMSYS','EXFSYS','ORDSYS','MDSYS', 'XDB', 'OUTLN', 'CTXMGR', 'OEMGR')
            owner = owner.ToUpper();

            switch (owner)
            {
                case "SYS":
                case "SYSMAN":
                case "SYSTEM":
                case "WMSYS":
                case "EXFSYS":
                case "ORDSYS":
                case "MDSYS":
                case "XDB":
                case "OUTLN":
                case "CTXMGR":
                case "OEMGR":
                    return true;

                default:
                    return false;
            }
        }


        public override bool CanEditConnectionString
        {
            get
            {
                return true;
            }
        }

        public override string EditConnectionString(string connectionString)
        {
            using (OracleConnectionEditor frm = new OracleConnectionEditor(connectionString))
            {
                if (frm.ShowDialog() == DialogResult.OK)
                    return frm.OracleBuilder.ToString();
            }
            return connectionString;
        }

        public override DialogResult Edit(bool newDatabase)
        {
            using (StiSqlDatabaseEditForm form = new StiSqlDatabaseEditForm(this))
            {
                if (newDatabase) form.Text = StiLocalization.Get("FormDatabaseEdit", "OracleODPNew");
                else form.Text = StiLocalization.Get("FormDatabaseEdit", "OracleODPEdit");

                form.tbName.Text = this.Name;
                form.tbAlias.Text = this.Alias;
                form.tbConnectionString.Text = this.ConnectionString;
                if (form.ShowDialog() == DialogResult.OK)
                {
                    this.Name = form.tbName.Text;
                    this.Alias = form.tbAlias.Text;
                    this.ConnectionString = form.tbConnectionString.Text;
                }
                return form.DialogResult;
            }
        }


        /// <summary>
        /// Creates a new object of the type StiOracleODPDatabase.
        /// </summary>
        public StiOracleODPDatabase()
            : this(string.Empty, string.Empty)
        {
        }


        /// <summary>
        /// Creates a new object of the type StiOracleODPDatabase.
        /// </summary>
        public StiOracleODPDatabase(string name, string connectionString)
            : base(name, connectionString)
        {
        }


        /// <summary>
        /// Creates a new object of the type StiOracleODPDatabase.
        /// </summary>
        public StiOracleODPDatabase(string name, string alias, string connectionString)
            : base(name, alias, connectionString)
        {
        }


        /// <summary>
        /// Creates a new object of the type StiOracleODPDatabase.
        /// </summary>
        public StiOracleODPDatabase(string name, string alias, string connectionString, bool promptUserNameAndpassword)
            :
            base(name, alias, connectionString, promptUserNameAndpassword)
        {
        }

    }

}
