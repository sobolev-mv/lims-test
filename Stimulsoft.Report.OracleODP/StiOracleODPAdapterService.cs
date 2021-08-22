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
using Oracle.DataAccess.Client;
using System.Windows.Forms;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Dictionary.Design;
using Stimulsoft.Base;
using Stimulsoft.Base.Localization;

namespace Stimulsoft.Report.Dictionary
{

    public class StiOracleODPAdapterService : StiSqlAdapterService
    {
        #region StiService override
        public override string ServiceName
        {
            get
            {
                return StiLocalization.Get("Adapters", "AdapterOracleODPConnection");
            }
        }
        #endregion

        #region StiDataAdapterService override
        public override StiDataColumnsCollection GetColumnsFromData(StiData data, StiDataSource dataSource)
        {
            StiDataColumnsCollection dataColumns = new StiDataColumnsCollection();
            StiOracleODPSource sqlSource = dataSource as StiOracleODPSource;

            try
            {
                if (sqlSource.SqlCommand != null && sqlSource.SqlCommand.Length > 0)
                {
                    if (data.Data is OracleConnection)
                    {
                        OracleConnection connection = data.Data as OracleConnection;
                        OpenConnection(connection, data, dataSource.Dictionary);

                        bool storedProc = false;
                        if (sqlSource.Type == StiSqlSourceType.StoredProcedure)
                        {
                            storedProc = true;
                        }
                        #region Stored Procedure
                        if (storedProc)
                        {
                            try
                            {
                                DataTable dataTable = new DataTable();
                                dataTable.TableName = sqlSource.Name;

                                OracleCommand myCommand = new OracleCommand(sqlSource.Name, connection);
                                myCommand.CommandType = CommandType.StoredProcedure;
                                myCommand.CommandTimeout = sqlSource.CommandTimeout;

                                //OracleCommandBuilder.DeriveParameters(myCommand);

                                foreach (StiDataParameter param in sqlSource.Parameters)
                                {
                                    OracleParameter oraParam = new OracleParameter(param.Name, (OracleDbType)param.Type);
                                    if ((OracleDbType)param.Type == OracleDbType.RefCursor)
                                        oraParam.Direction = ParameterDirection.Output;
                                    else
                                        oraParam.Direction = ParameterDirection.Input;

                                    myCommand.Parameters.Add(oraParam);
                                }

                                //dataTable.Load(reader);
                                OracleDataReader reader = myCommand.ExecuteReader(CommandBehavior.CloseConnection);
                                dataTable.Load(reader);

                                foreach (DataColumn column in dataTable.Columns)
                                {
                                    dataColumns.Add(new StiDataColumn(column.ColumnName, column.Caption, column.DataType));
                                }
                                dataTable.Dispose();
                            }
                            catch
                            {
                                storedProc = false;
                            }
                        }
                        #endregion
                        if (!storedProc)
                        {
                            using (OracleDataAdapter dataAdapter = new OracleDataAdapter(sqlSource.SqlCommand, connection))
                            {
                                DataTable dataTable = new DataTable();
                                dataTable.TableName = sqlSource.Name;

                                dataAdapter.FillSchema(dataTable, SchemaType.Source);

                                foreach (DataColumn column in dataTable.Columns)
                                {
                                    dataColumns.Add(new StiDataColumn(column.ColumnName, column.Caption, column.DataType));
                                }
                                dataTable.Dispose();
                            }
                        }
                        CloseConnection(data, connection);
                    }
                }
            }
            catch (Exception e)
            {
                StiLogService.Write(this.GetType(), e);
                if (!StiOptions.Engine.HideExceptions) throw;
            }

            return dataColumns;
        }

        public override void SetDataSourceNames(StiData data, StiDataSource dataSource)
        {
            base.SetDataSourceNames(data, dataSource);

            StiDataColumnsCollection dataColumns = new StiDataColumnsCollection();
            StiSqlSource sqlSource = dataSource as StiSqlSource;

            dataSource.Name = "OracleODPSource";
            dataSource.Alias = "OracleODPSource";
        }

        public override Type GetDataSourceType()
        {
            return typeof(StiOracleODPSource);
        }

        public override Type[] GetDataTypes()
        {
            return new Type[] { typeof(OracleConnection) };
        }

        public override void ConnectDataSourceToData(StiDictionary dictionary, StiDataSource dataSource, bool loadData)
        {
            dataSource.Disconnect();

            if (!loadData)
            {
                dataSource.DataTable = new DataTable();
                return;
            }

            StiSqlSource sqlSource = dataSource as StiSqlSource;

            foreach (StiData data in dataSource.Dictionary.DataStore)
            {
                if (data.Name == sqlSource.NameInSource)
                {
                    try
                    {
                        if (data.Data is OracleConnection)
                        {
                            OracleConnection connection = data.ViewData as OracleConnection;
                            OpenConnection(connection, data, dataSource.Dictionary);

                            sqlSource.DataAdapter = new OracleDataAdapter(sqlSource.SqlCommand, connection);

                            foreach (StiDataParameter parameter in sqlSource.Parameters)
                            {
                                ((OracleDataAdapter)sqlSource.DataAdapter).SelectCommand.Parameters.Add(
                                    parameter.Name, (OracleDbType)parameter.Type, parameter.Size);
                            }

                            DataTable dataTable = new DataTable();
                            dataTable.TableName = sqlSource.Name;
                            dataSource.DataTable = dataTable;


                            sqlSource.DataAdapter.SelectCommand.CommandTimeout = sqlSource.CommandTimeout;

                            if (loadData && sqlSource.Parameters.Count > 0)
                            {
                                sqlSource.DataAdapter.SelectCommand.Prepare();
                                sqlSource.UpdateParameters();
                            }
                            else
                            {
                                if (loadData)
                                {
                                    ((OracleDataAdapter)sqlSource.DataAdapter).Fill(dataTable);
                                    sqlSource.CheckColumnsIndexs();
                                }
                                else ((OracleDataAdapter)sqlSource.DataAdapter).FillSchema(dataTable, SchemaType.Source);
                            }

                            break;
                        }
                    }
                    catch (Exception e)
                    {
                        StiLogService.Write(this.GetType(), e);
                        if (!StiOptions.Engine.HideExceptions) throw;
                    }
                }
            }
        }
        #endregion

        #region StiSqlAdapterService override
        public override void CreateConnectionInDataStore(StiDictionary dictionary, StiSqlDatabase database)
        {
            try
            {
                #region remove all old data from datastore
                int index = 0;
                foreach (StiData data in dictionary.DataStore)
                {
                    if (data.Name == database.Name)
                    {
                        dictionary.DataStore.RemoveAt(index);
                        break;
                    }
                    index++;
                }
                #endregion

                OracleConnection sqlConnection = new OracleConnection(database.ConnectionString);
                StiData data2 = new StiData(database.Name, sqlConnection);
                data2.IsReportData = true;
                dictionary.DataStore.Add(data2);

            }
            catch (Exception e)
            {
                StiLogService.Write(this.GetType(), e);
                if (!StiOptions.Engine.HideExceptions) throw;
            }
        }

        public override string TestConnection(string connectionString)
        {
            try
            {
                using (OracleConnection sqlConnection = new OracleConnection(connectionString))
                {
                    sqlConnection.Open();
                    sqlConnection.Close();
                    return StiLocalization.Get("DesignerFx", "ConnectionSuccessfull");
                }
            }
            catch (Exception e)
            {
                return StiLocalization.Get("DesignerFx", "ConnectionError") + ": " + e.Message;
            }
        }
        #endregion

        #region GetQueryBuilderProviders
        public override object GetSyntaxProvider()
        {
            return new Stimulsoft.Database.SQL92SyntaxProvider();
        }
        public override object GetMetadataProvider(IDbConnection connection)
        {
            return new Stimulsoft.Database.StiOracleODPMetadataProvider(connection);
        }
        #endregion

    }
}
