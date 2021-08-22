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

#if Net2
using System;
using System.ComponentModel;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Reflection;
using System.Reflection.Emit;
using Oracle.DataAccess.Client;
using Stimulsoft.Base;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Base.Localization;
using Stimulsoft.Base.Services;

namespace Stimulsoft.Database
{
    [StiServiceBitmap(typeof(Stimulsoft.Report.Dictionary.StiDataAdapterService), "Stimulsoft.Report.Bmp.DataAdapter.bmp")]
    [StiServiceCategoryBitmap(typeof(Stimulsoft.Report.Dictionary.StiDataAdapterService), "Stimulsoft.Report.Bmp.DataAdapter.bmp")]		
    public class StiOracleODPMetadataProvider : BaseMetadataProvider
    {
        #region StiService override
        /// <summary>
        /// Gets a service name.
        /// </summary>
        public override string ServiceName
        {
            get
            {
                return "Oracle Metadata Provider";
            }
        }
        #endregion

        #region Properties
        [Browsable(true)]
        public OracleConnection Connection
        { 
            get 
            { 
                return internalConnection as OracleConnection; 
            } 
            set 
            {
                if (internalConnection != value)
                {
                    internalConnection = value;
                }
            }
        }        

        public override bool IsConnected
        {
            get
            {
                if (Connection != null)
                {
                    return (Connection.State & ConnectionState.Open) == ConnectionState.Open;
                }
                else return false;
            }
            set
            {
                base.IsConnected = value;
            }
        }

        public override bool IsCanCreateInternalConnection
        {
            get
            {
                return true;
            }
        }
        #endregion

        #region Methods
        public override void ViewQuery(
            IDbConnection connection, 
            QueryBuilder queryBuilder, DataGridView dataGridView)
        {
            using (OracleCommand cmd = connection.CreateCommand() as OracleCommand)
            {
                cmd.CommandTimeout = Stimulsoft.Report.StiOptions.Dictionary.QueryBuilderConnectTimeout;
                cmd.CommandText = queryBuilder.SQL;

                // handle the query parameters
                if (queryBuilder.Parameters.Count > 0)
                {
                    for (int i = 0; i < queryBuilder.Parameters.Count; i++)
                    {
                        OracleParameter p = new OracleParameter();
                        p.ParameterName = queryBuilder.Parameters[i].FullName;
                        p.DbType = queryBuilder.Parameters[i].DataType;
                        cmd.Parameters.Add(p);
                    }

                    using (QueryParametersForm form = new QueryParametersForm(queryBuilder.Parameters, cmd))
                    {
                        form.ShowDialog();
                    }
                }

                using (OracleDataAdapter adapter = new OracleDataAdapter(cmd))
                {
                    DataSet dataset = new DataSet();

                    try
                    {
                        adapter.Fill(dataset, "QueryResult");
                        dataGridView.DataSource = dataset.Tables["QueryResult"];
                    }
                    catch (Exception ex)
                    {
                        Stimulsoft.Base.StiExceptionProvider.Show(ex);
                    }
                }
            }
        }

        protected override void DoConnect()
        {
            base.DoConnect();

            CheckConnectionSet();

            try
            {
                Connection.Open();
            }
            catch
            {
                throw;
            }
        }

        protected override void DoDisconnect()
        {
            base.DoDisconnect();

            CheckConnectionSet();
            Connection.Close();
        }

        protected override void CheckConnectionSet()
        {
            if (Connection == null)
            {
                throw new Exception(String.Format(StiLocalization.Get("QueryBuilder", "NoConnectionObject"), "Connection"));
            }
        }

        protected override IDataReader PrepareSQLDatasetInternal(string sql, bool schemaOnly)
        {
            OracleCommand command = null;
            OracleDataReader reader = null;

            if (!IsConnected) Connect();

            command = Connection.CreateCommand();
            command.CommandText = sql;

            if (schemaOnly)
            {
                reader = command.ExecuteReader(CommandBehavior.SchemaOnly);
            }
            else
            {
                reader = command.ExecuteReader();
            }

            return reader;
        }

        protected override void ExecSQLInternal(string sql)
        {
            if (!IsConnected) Connect();

            OracleCommand command = Connection.CreateCommand();
            command.CommandText = sql;
            command.ExecuteNonQuery();
        }

        public override void LoadMetadataObjects(
            BaseSyntaxProvider ASyntaxProvider, 
            MetadataFilter AMetadataFilter, 
            MetadataContainer AMetadataContainer, 
            SQLQualifiedName ADatabase)
        {
            MetadataObjectFetcherFromQuery mof;

            if (IsCanExecSQL && Helpers.IsQualifiedNameEmpty(ADatabase))
            {
                mof = new MetadataObjectFetcherFromQuery(ASyntaxProvider, this);

                mof.Query = "select TABLE_NAME, Owner from all_tables where " + Helpers.strTableFilterMacro;
                mof.SchemaFieldName = "OWNER";
                mof.NameFieldName = "TABLE_NAME";
                mof.DefaultObjectClass = typeof(MetadataTable);
                mof.SystemSchemaNames.AddIdentifier("SYS");
                mof.SystemSchemaNames.AddIdentifier("SYSTEM");
                mof.SystemSchemaNames.AddIdentifier("OUTLN");
                mof.SystemSchemaNames.AddIdentifier("WMSYS");
                mof.SystemSchemaNames.AddIdentifier("CTXSYS");
                mof.SystemSchemaNames.AddIdentifier("ORDSYS");
                mof.SystemSchemaNames.AddIdentifier("XDB");
                mof.SystemSchemaNames.AddIdentifier("MDSYS");
                mof.SystemSchemaNames.AddIdentifier("TSMSYS");
                mof.SystemSchemaNames.AddIdentifier("LBACSYS");

                mof.LoadMetadata(AMetadataFilter, AMetadataContainer, ADatabase);

                mof = new MetadataObjectFetcherFromQuery(ASyntaxProvider, this);

                mof.Query = "select VIEW_NAME, Owner from all_views where " + Helpers.strViewFilterMacro;
                mof.SchemaFieldName = "OWNER";
                mof.NameFieldName = "VIEW_NAME";
                mof.DefaultObjectClass = typeof(MetadataView);
                mof.SystemSchemaNames.AddIdentifier("SYS");
                mof.SystemSchemaNames.AddIdentifier("SYSTEM");
                mof.SystemSchemaNames.AddIdentifier("OUTLN");
                mof.SystemSchemaNames.AddIdentifier("WMSYS");
                mof.SystemSchemaNames.AddIdentifier("CTXSYS");
                mof.SystemSchemaNames.AddIdentifier("ORDSYS");
                mof.SystemSchemaNames.AddIdentifier("XDB");
                mof.SystemSchemaNames.AddIdentifier("MDSYS");
                mof.SystemSchemaNames.AddIdentifier("TSMSYS");
                mof.SystemSchemaNames.AddIdentifier("LBACSYS");

                mof.LoadMetadata(AMetadataFilter, AMetadataContainer, ADatabase);
            }
            else
            {
                base.LoadMetadataObjects(ASyntaxProvider, AMetadataFilter, AMetadataContainer, ADatabase);
            }
        }

        public override void LoadMetadataRelations(
            BaseSyntaxProvider ASyntaxProvider, 
            MetadataFilter AMetadataFilter, 
            MetadataContainer AMetadataContainer, 
            SQLQualifiedName ADatabase)
        {
            if (IsCanExecSQL)
            {
                ConstraintsList cl = new ConstraintsList();
                Constraint c;
                ConstraintColumn cc;
                string s;

                // load constraint into list

                string sql = "select owner, constraint_name, constraint_type, r_owner, r_constraint_name from all_constraints where (constraint_type in ('R','P','U'))";
                int OwnerField, NameField, TypeField, ROwnerField, RNameField;

                IDataReader reader = this.ExecSQL(sql, false);

                OwnerField = reader.GetOrdinal("Owner");
                NameField = reader.GetOrdinal("constraint_name");
                TypeField = reader.GetOrdinal("constraint_type");
                ROwnerField = reader.GetOrdinal("r_owner");
                RNameField = reader.GetOrdinal("r_constraint_name");

                while (reader.Read())
                {
                    s = reader.GetString(TypeField);

                    if (s == "R" || s == "P" || s == "U")
                    {
                        c = new Constraint();
                        cl.Add(c);
                        c.Owner = new AstTokenIdentifier(ASyntaxProvider, reader.GetString(OwnerField), true);
                        c.Name = new AstTokenIdentifier(ASyntaxProvider, reader.GetString(NameField), true);
                        c.ConstraintType = s;
                        c.ROwner = new AstTokenIdentifier(ASyntaxProvider, (reader.IsDBNull(ROwnerField)) ? "" : reader.GetString(ROwnerField), true);
                        c.RName = new AstTokenIdentifier(ASyntaxProvider, (reader.IsDBNull(RNameField)) ? "" : reader.GetString(RNameField), true);
                    }
                }

                // load constraints columns into list

                int TableField, ColumnField;
                AstTokenIdentifier ownerId, nameId;
                sql = "select owner, constraint_name, table_name, column_name from all_cons_columns order by owner, constraint_name, table_name, position";

                reader = this.ExecSQL(sql, false);

                OwnerField = reader.GetOrdinal("owner");
                NameField = reader.GetOrdinal("constraint_name");
                TableField = reader.GetOrdinal("table_name");
                ColumnField = reader.GetOrdinal("column_name");

                while (reader.Read())
                {
                    ownerId = new AstTokenIdentifier(ASyntaxProvider, reader.GetString(OwnerField), true);
                    nameId = new AstTokenIdentifier(ASyntaxProvider, reader.GetString(NameField), true);
                    c = cl.FindConstraint(ownerId, nameId);

                    if (c != null)
                    {
                        cc = new ConstraintColumn();
                        c.Columns.Add(cc);
                        cc.Owner = ownerId;
                        cc.Name = nameId;
                        cc.TableName = new AstTokenIdentifier(ASyntaxProvider, reader.GetString(TableField), true);
                        cc.ColName = new AstTokenIdentifier(ASyntaxProvider, reader.GetString(ColumnField), true);
                    }
                }

                // load constraints

                Constraint childConstraint, parentConstraint;
                MetadataObject pt;
                MetadataRelation r;

                for (int iRelations = 0; iRelations < cl.Count; iRelations++)
                {
                    childConstraint = (Constraint)cl[iRelations];

                    // for all reference integrity constraints
                    if (childConstraint.ConstraintType == "R" && childConstraint.ROwner != null && childConstraint.ROwner.Token != "" &&
                        childConstraint.RName != null && childConstraint.RName.Token != "")
                    {
                        // find parent constraint
                        parentConstraint = cl.FindConstraint(childConstraint.ROwner, childConstraint.RName);

                        if (parentConstraint != null && childConstraint.Columns.Count > 0 &&
                            childConstraint.Columns.Count == parentConstraint.Columns.Count)
                        {
                            // find parent table
                            pt = AMetadataContainer.FindObjectByName(((ConstraintColumn)parentConstraint.Columns[0]).TableName,
                                ((ConstraintColumn)parentConstraint.Columns[0]).Owner, null);

                            // create relation
                            r = pt.Relations.Add();
                            r.ChildName = ((ConstraintColumn)childConstraint.Columns[0]).TableName;
                            r.ChildSchema = ((ConstraintColumn)childConstraint.Columns[0]).Owner;

                            for (int iColumns = 0; iColumns < parentConstraint.Columns.Count; iColumns++)
                            {
                                r.KeyFields.AddField(((ConstraintColumn)parentConstraint.Columns[iColumns]).ColName);
                                r.ChildFields.AddField(((ConstraintColumn)childConstraint.Columns[iColumns]).ColName);
                            }

                            // check unique
                            int i = r.Relations.FindRelation(r.KeyFields, r.ChildSchema, r.ChildName, r.ChildFields, null);

                            if (i < r.Relations.Count - 1)
                            {
                                r.Relations.Remove(r);
                            }
                        }
                    }
                }
            }
            else
            {
                base.LoadMetadataRelations(ASyntaxProvider, AMetadataFilter, AMetadataContainer, ADatabase);
            }
        }

        internal class Constraint
        {
            private ConstraintsColumnsList fColumns = new ConstraintsColumnsList();
            public AstTokenIdentifier Owner = null;
            public AstTokenIdentifier Name = null;
            public string ConstraintType = null;
            public AstTokenIdentifier ROwner = null;
            public AstTokenIdentifier RName = null;

            public ConstraintsColumnsList Columns { get { return fColumns; } }
        }

        internal class ConstraintsList : ArrayList
        {
            public ConstraintsList() { }

            public Constraint FindConstraint(AstTokenIdentifier AOwner, AstTokenIdentifier AName)
            {
                Constraint c;

                Debug.Assert(AOwner != null);
                Debug.Assert(AName != null);

                for (int i = 0; i < Count; i++)
                {
                    c = (Constraint)this[i];

                    if (AName.SyntaxProvider.IsIdentifiersEqual(c.Name, AName) && AName.SyntaxProvider.IsIdentifiersEqual(c.Owner, AOwner))
                    {
                        return c;
                    }
                }

                return null;
            }
        }

        internal class ConstraintColumn
        {
            public AstTokenIdentifier Owner = null;
            public AstTokenIdentifier Name = null;
            public AstTokenIdentifier TableName = null;
            public AstTokenIdentifier ColName = null;
        }

        internal class ConstraintsColumnsList : ArrayList
        {
            public ConstraintsColumnsList() { }
        }

        
        public override void CreateAndBindInternalConnectionObj()
        {
            base.CreateAndBindInternalConnectionObj();
            Connection = new OracleConnection();
            AddInternalConnectionObject(Connection);
        }
        #endregion

        public StiOracleODPMetadataProvider()
            : this(null)
        {
        }

        public StiOracleODPMetadataProvider(IDbConnection connection) : base(connection)
        {
        }

        static StiOracleODPMetadataProvider()
        {
            Stimulsoft.Database.Helpers.MetadataProviderList.RegisterMetadataProvider(typeof(StiOracleODPMetadataProvider));
        }
    }
}
#endif

