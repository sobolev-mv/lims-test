using System;
using System.Data;
using System.Collections.Generic;
using Smv.Data.Oracle;
using Devart.Data.Oracle;

namespace Viz.WrkModule.Thp.Db.DataSets
{
  public sealed class DsThp : DataSet
  {
    public ThpDataTable Thp { get; private set; }
    public ThpDetailDataTable ThpDetail { get; private set; }

    public DsThp() : base()
    {
      this.DataSetName = "DsThp";

      this.Thp = new ThpDataTable("Thp");
      this.Tables.Add(this.Thp);

      this.ThpDetail = new ThpDetailDataTable("ThpDetail");
      this.Tables.Add(this.ThpDetail);

    }

    public class ThpDataTable : DataTable
    {
      protected readonly OracleDataAdapter adapter;

      public ThpDataTable(string tblName) : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof (Int64), null, MappingType.Element)
        {
          AllowDBNull = false,
          AutoIncrement = true,
          AutoIncrementStep = -1
        };
        this.Columns.Add(col);

        col = new DataColumn("Pid", typeof (Int64), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("CodeThp", typeof (string), null, MappingType.Element) {AllowDBNull = false};
        this.Columns.Add(col);

        col = new DataColumn("NumThp", typeof (string), null, MappingType.Element) {AllowDBNull = false};
        this.Columns.Add(col);

        col = new DataColumn("TypeThp", typeof (string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("DateThp", typeof (DateTime), null, MappingType.Element) {AllowDBNull = false};
        this.Columns.Add(col);

        col = new DataColumn("DateBeg", typeof (DateTime), null, MappingType.Element) {AllowDBNull = false};
        this.Columns.Add(col);

        col = new DataColumn("DateEnd", typeof (DateTime), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Ts", typeof (string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("SubjThp", typeof (string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("NoteThp", typeof (string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("IsCancel", typeof (Int32), null, MappingType.Element) {DefaultValue = 0};
        this.Columns.Add(col);

        col = new DataColumn("TypeDoc", typeof (string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("TypeProt", typeof (string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Ldoc", typeof (Int32), null, MappingType.Element) {DefaultValue = 0};
        this.Columns.Add(col);

        col = new DataColumn("Lprot", typeof (Int32), null, MappingType.Element) {DefaultValue = 0};
        this.Columns.Add(col);

        col = new DataColumn("Cdtl", typeof(Int32), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] {this.Columns["Id"]}, true));
        this.Columns["Id"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("LIMS.THP_DATA", tblName);
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("PID", "Pid");
        dtm.ColumnMappings.Add("CODE_THP", "CodeThp");
        dtm.ColumnMappings.Add("NUM_THP", "NumThp");
        dtm.ColumnMappings.Add("TYPE_TXP", "TypeThp");
        dtm.ColumnMappings.Add("DATE_THP", "DateThp");
        dtm.ColumnMappings.Add("DATE_BEG", "DateBeg");
        dtm.ColumnMappings.Add("DATE_END", "DateEnd");
        dtm.ColumnMappings.Add("TS", "Ts");
        dtm.ColumnMappings.Add("SUBJ_THP", "SubjThp");
        dtm.ColumnMappings.Add("NOTE_THP", "NoteThp");
        dtm.ColumnMappings.Add("IS_CANCEL", "IsCancel");
        dtm.ColumnMappings.Add("TYPE_DOC", "TypeDoc");
        dtm.ColumnMappings.Add("TYPE_PROT", "TypeProt");
        dtm.ColumnMappings.Add("L_DOC", "Ldoc");
        dtm.ColumnMappings.Add("L_PROT", "Lprot");
        dtm.ColumnMappings.Add("CDTL", "Cdtl");

        adapter.TableMappings.Add(dtm);

        //Select Command
        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "SELECT ID, PID, CODE_THP, NUM_THP, TYPE_TXP, DATE_THP, DATE_BEG, DATE_END, TS, SUBJ_THP, NOTE_THP, IS_CANCEL, TYPE_DOC, TYPE_PROT, L_DOC, L_PROT, CDTL " +
            "FROM LIMS.V_THPDATA WHERE (DATE_THP BETWEEN :DT1 AND :DT2) AND (PID IS NULL) ORDER BY DATE_THP",
          CommandType = CommandType.Text
        };

        var param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "DT1",
          SourceColumn = "DATE_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "DT2",
          SourceColumn = "DATE_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        //Insert Command
        adapter.InsertCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "INSERT INTO LIMS.THP_DATA(PID, CODE_THP, NUM_THP, TYPE_TXP, DATE_THP, DATE_BEG, DATE_END, TS, SUBJ_THP, NOTE_THP, IS_CANCEL) " +
            "VALUES(:PPID, :PCODE_THP, :PNUM_THP, :PTYPE_TXP, :PDATE_THP, :PDATE_BEG, :PDATE_END, :PTS, :PSUBJ_THP, :PNOTE_THP, :PIS_CANCEL) RETURNING ID, IS_CANCEL INTO :PID, :POIS_CANCEL",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.OutputParameters
        };

        param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PPID",
          SourceColumn = "PID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PCODE_THP",
          SourceColumn = "CODE_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PNUM_THP",
          SourceColumn = "NUM_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTYPE_TXP",
          SourceColumn = "TYPE_TXP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_THP",
          SourceColumn = "DATE_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_BEG",
          SourceColumn = "DATE_BEG",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_END",
          SourceColumn = "DATE_END",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTS",
          SourceColumn = "TS",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PSUBJ_THP",
          SourceColumn = "SUBJ_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PNOTE_THP",
          SourceColumn = "NOTE_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "PIS_CANCEL",
          SourceColumn = "IS_CANCEL",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.ReturnValue,
          ParameterName = "PID",
          SourceColumn = "ID",
          SourceColumnNullMapping = false,
          Precision = 15,
          Scale = 0,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.ReturnValue,
          ParameterName = "POIS_CANCEL",
          SourceColumn = "IS_CANCEL",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        //Update Command
        adapter.UpdateCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "UPDATE LIMS.THP_DATA SET CODE_THP = :PCODE_THP,  NUM_THP = :PNUM_THP, TYPE_TXP = :PTYPE_TXP, DATE_THP = :PDATE_THP, DATE_BEG = :PDATE_BEG, " +
            "DATE_END = :PDATE_END, TS = :PTS, SUBJ_THP = :PSUBJ_THP, NOTE_THP = :PNOTE_THP, IS_CANCEL = :PIS_CANCEL WHERE (ID = :Original_ID)",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.None
        };

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PCODE_THP",
          SourceColumn = "CODE_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PNUM_THP",
          SourceColumn = "NUM_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTYPE_TXP",
          SourceColumn = "TYPE_TXP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_THP",
          SourceColumn = "DATE_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_BEG",
          SourceColumn = "DATE_BEG",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_END",
          SourceColumn = "DATE_END",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTS",
          SourceColumn = "TS",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PSUBJ_THP",
          SourceColumn = "SUBJ_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PNOTE_THP",
          SourceColumn = "NOTE_THP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "PIS_CANCEL",
          SourceColumn = "IS_CANCEL",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "Original_ID",
          SourceColumn = "ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Original
        };
        adapter.UpdateCommand.Parameters.Add(param);
      }

      public int LoadData(DateTime dtBegin, DateTime dtEnd)
      {
        var lstPrmValue = new List<Object> {dtBegin, dtEnd};
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }


      public int SaveData()
      {
        return Odac.SaveChangedData(this, adapter);
      }

    }

    public sealed class ThpDetailDataTable : ThpDataTable
    {
      public ThpDetailDataTable(string tblName) : base(tblName)
      {
        //Select Command
        adapter.SelectCommand.Parameters.Clear();

        adapter.SelectCommand.CommandText =
          "SELECT ID, PID, CODE_THP, NUM_THP, TYPE_TXP, DATE_THP, DATE_BEG, DATE_END, TS, SUBJ_THP, NOTE_THP, IS_CANCEL, TYPE_DOC, TYPE_PROT, L_DOC, L_PROT, CDTL " +
          "FROM LIMS.V_THPDATA WHERE PID = :PPID ORDER BY DATE_THP";

        var param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PPID",
          SourceColumn = "PID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);
      }

      public int LoadData(Int64 parentId)
      {
        var lstPrmValue = new List<Object> {parentId};
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }

    }
  }
}
