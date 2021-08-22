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
#endregion Copyright (C) 2003-2011 Stimulsoft

using System;
using System.Collections;
using System.Data;
using System.ComponentModel;
using Oracle.DataAccess.Client;
using Stimulsoft.Base;
using Stimulsoft.Base.Localization;
using Stimulsoft.Base.Serializing;
using Stimulsoft.Report.Dictionary.Design;
using Stimulsoft.Report.Events;

namespace Stimulsoft.Report.Dictionary
{

    [TypeConverter(typeof(StiSqlSourceConverter))]
    public class StiOracleODPSource : StiSqlSource
    {
        #region DataAdapter
        protected override Type ConvertDbTypeToTypeInternal(int sqlType)
        {
            OracleDbType dbType = (OracleDbType)sqlType;
            switch (dbType)
            {
                case OracleDbType.Byte:
                case OracleDbType.Int16:
                case OracleDbType.Int32:
                    //case OracleDbType.RowId:
                    //case OracleDbType.UInt16:
                    //case OracleDbType.UInt32:
                    return typeof(Int64);

                case OracleDbType.Decimal:
                    return typeof(decimal);

                case OracleDbType.Double:
                    return typeof(double);

                case OracleDbType.Date:
                case OracleDbType.TimeStamp:
                case OracleDbType.TimeStampTZ:
                case OracleDbType.TimeStampLTZ:
                    return typeof(DateTime);

                default:
                    return typeof(string);
            }
        }

        public override Type GetParameterTypesEnum()
        {
            return typeof(OracleDbType);
        }

        public override StiDataParameter AddParameter()
        {
            StiDataParameter parameter = new StiDataParameter(
                StiLocalization.Get("PropertyMain", "Parameter"), string.Empty, (int)OracleDbType.Varchar2, 0);
            Parameters.Add(parameter);
            return parameter;
        }

        public virtual StiDataParameter AddParameter(string name, string expression, OracleDbType type, int size)
        {
            StiDataParameter parameter = AddParameter();
            parameter.Name = name;
            parameter.Expression = expression;
            parameter.Type = (int)type;
            parameter.Size = size;

            return parameter;
        }

        public override void UpdateParameters()
        {
            if (this.DataTable != null)
            {
                InvokeConnecting();
                foreach (StiDataParameter parameter in Parameters)
                {
                    ((OracleDataAdapter)DataAdapter).SelectCommand.Parameters[parameter.Name].Value = parameter.GetParameterValue();

                    if (((OracleDataAdapter)DataAdapter).SelectCommand.Parameters[parameter.Name].OracleDbType == OracleDbType.RefCursor)
                        ((OracleDataAdapter)DataAdapter).SelectCommand.Parameters[parameter.Name].Direction = ParameterDirection.Output;
                }

                DataTable dataTable = this.DataTable;
                dataTable.Rows.Clear();

                if (this.Type == StiSqlSourceType.Table)
                {

                    ((OracleDataAdapter)DataAdapter).SelectCommand.CommandTimeout = this.CommandTimeout;
                    ((OracleDataAdapter)DataAdapter).Fill(dataTable);
                }
                else
                {
                    ((OracleDataAdapter)DataAdapter).SelectCommand.CommandType = CommandType.StoredProcedure;
                    OracleDataReader reader = ((OracleDataAdapter)DataAdapter).SelectCommand.ExecuteReader(CommandBehavior.CloseConnection);
                    dataTable.Load(reader);
                }
                CheckColumnsIndexs();
            }
        }

        protected override string DataAdapterType
        {
            get
            {
                return "Stimulsoft.Report.Dictionary.StiOracleODPAdapterService";
            }
        }
        #endregion

        #region this
        public StiOracleODPSource()
            : this("", "", "")
        {
        }

        public StiOracleODPSource(string nameInSource, string name)
            : this(nameInSource, name, name)
        {
        }

        public StiOracleODPSource(string nameInSource, string name, string alias)
            : this(nameInSource, name, alias, string.Empty)
        {

        }

        public StiOracleODPSource(string nameInSource, string name, string alias, string sqlCommand)
            :
            base(nameInSource, name, alias, sqlCommand)
        {
        }

        public StiOracleODPSource(string nameInSource, string name, string alias, string sqlCommand,
            bool connectOnStart)
            :
        base(nameInSource, name, alias, sqlCommand, connectOnStart)
        {
        }

        public StiOracleODPSource(string nameInSource, string name, string alias, string sqlCommand,
            bool connectOnStart, bool reconnectOnEachRow)
            :
        base(nameInSource, name, alias, sqlCommand, connectOnStart, reconnectOnEachRow)
        {
        }

        public StiOracleODPSource(string nameInSource, string name, string alias, string sqlCommand,
            bool connectOnStart, bool reconnectOnEachRow, int commandTimeout)
            :
            base(nameInSource, name, alias, sqlCommand, connectOnStart, reconnectOnEachRow, commandTimeout)
        {
        }
        #endregion
    }

}


