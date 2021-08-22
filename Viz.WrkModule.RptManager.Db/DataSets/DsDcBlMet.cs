using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.Common;
using System.Data;
using Smv.Data;
using Smv.Data.Oracle;
using Devart.Data.Oracle;

namespace Viz.WrkModule.RptManager.Db.DataSets
{
  public sealed class DsDcBlMet : DataSet
  {
    public DcBlMetDataTable DcBlMet { get; private set; }
    public ParamListDataTable ParamListThickness { get; private set; }
    public ThicknessDataTable Thickness { get; private set; }

    public DsDcBlMet()
    {
      DataSetName = "DsDcBlMet";

      DcBlMet = new DcBlMetDataTable("DcBlMet");
      Tables.Add(DcBlMet);

      ParamListThickness = new ParamListDataTable("ParamListThickness");
      Tables.Add(ParamListThickness);

      Thickness = new ThicknessDataTable("Thickness");
      Tables.Add(Thickness);
    }

    public sealed class DcBlMetDataTable : SmvDataTable
    {
      private readonly OracleDataAdapter adapter;

      public DcBlMetDataTable(string tblName)
      {
        TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof(Int32), null, MappingType.Element)
        {
          AllowDBNull = false,
          AutoIncrement = true,
          AutoIncrementSeed = -1,
          AutoIncrementStep = -1
        };
        Columns.Add(col);

        col = new DataColumn("DateFrom", typeof(DateTime), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("DateTo", typeof(DateTime), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("Pdc1StCut", typeof(decimal), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("Tdc1StCut", typeof(decimal), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("Pdc2NdCut", typeof(decimal), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("Tdc2NdCut", typeof(decimal), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("PdcStrann", typeof(decimal), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("TdcStrann", typeof(decimal), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("PdcUo", typeof(decimal), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("TdcUo", typeof(decimal), null, MappingType.Element) {AllowDBNull = false};
        Columns.Add(col);

        col = new DataColumn("IsLast", typeof(string), null, MappingType.Element) {MaxLength = 1};
        Columns.Add(col);

        col = new DataColumn("P1Sort", typeof(decimal), null, MappingType.Element);
        Columns.Add(col);

        Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] {Columns["Id"]}, true));
        Columns["Id"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new DataTableMapping("DG_DCBLMET", tblName);
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("DATE_FROM", "DateFrom");
        dtm.ColumnMappings.Add("DATE_TO", "DateTo");
        dtm.ColumnMappings.Add("PDC_1STCUT", "Pdc1StCut");
        dtm.ColumnMappings.Add("TDC_1STCUT", "Tdc1StCut");
        dtm.ColumnMappings.Add("PDC_2NDCUT", "Pdc2NdCut");
        dtm.ColumnMappings.Add("TDC_2NDCUT", "Tdc2NdCut");
        dtm.ColumnMappings.Add("PDC_STRANN", "PdcStrann");
        dtm.ColumnMappings.Add("TDC_STRANN", "TdcStrann");
        dtm.ColumnMappings.Add("PDC_UO", "PdcUo");
        dtm.ColumnMappings.Add("TDC_UO", "TdcUo");
        dtm.ColumnMappings.Add("IS_LAST", "IsLast");
        dtm.ColumnMappings.Add("P_1SORT", "P1Sort");
        adapter.TableMappings.Add(dtm);

        //Select Command
        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandType = CommandType.Text,
          UpdatedRowSource = UpdateRowSource.None
        };

        //Update Command
        adapter.UpdateCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "UPDATE VIZ_PRN.DG_DCBLMET SET PDC_1STCUT = :PDC_1STCUT, TDC_1STCUT = :TDC_1STCUT, PDC_2NDCUT = :PDC_2NDCUT, TDC_2NDCUT = :TDC_2NDCUT, " +
            "PDC_STRANN = :PDC_STRANN, TDC_STRANN = :TDC_STRANN, PDC_UO = :PDC_UO, TDC_UO = :TDC_UO, P_1SORT = :P_1SORT " +
            "WHERE (ID = :Original_ID)",
          CommandType = CommandType.Text,
          UpdatedRowSource = UpdateRowSource.None
        };

        var param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PDC_1STCUT",
          SourceColumn = "PDC_1STCUT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 6,
          Scale = 4
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "TDC_1STCUT",
          SourceColumn = "TDC_1STCUT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 6,
          Scale = 4
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PDC_2NDCUT",
          SourceColumn = "PDC_2NDCUT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 6,
          Scale = 4
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "TDC_2NDCUT",
          SourceColumn = "TDC_2NDCUT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 6,
          Scale = 4
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PDC_STRANN",
          SourceColumn = "PDC_STRANN",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 6,
          Scale = 4
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "TDC_STRANN",
          SourceColumn = "TDC_STRANN",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 6,
          Scale = 4
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "TDC_STRANN",
          SourceColumn = "TDC_STRANN",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 6,
          Scale = 4
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PDC_UO",
          SourceColumn = "PDC_UO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 6,
          Scale = 4
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "TDC_UO",
          SourceColumn = "TDC_UO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 6,
          Scale = 4
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "P_1SORT",
          SourceColumn = "P_1SORT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 4,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);


        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "Original_ID",
          SourceColumn = "ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Original
        };
        adapter.UpdateCommand.Parameters.Add(param);
      }

      public int LoadData()
      {
        adapter.SelectCommand.CommandText =
          "SELECT ID, DATE_FROM, DATE_TO, PDC_1STCUT, TDC_1STCUT, PDC_2NDCUT, TDC_2NDCUT, PDC_STRANN, TDC_STRANN, PDC_UO, TDC_UO, IS_LAST, P_1SORT " +
          "FROM VIZ_PRN.DG_DCBLMET ORDER BY ID DESC";

        adapter.SelectCommand.Parameters.Clear();
        return Odac.LoadDataTable(this, adapter, true, null);
      }

      public int SaveData()
      {
        return Odac.SaveChangedData(this, adapter);
      }
    }

    public sealed class ParamListDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public ParamListDataTable(string tblName) : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof(Int32), null, MappingType.Element) {AllowDBNull = false};
        this.Columns.Add(col);

        col = new DataColumn("StrSql", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("StrDlg", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] {this.Columns["Id"]}, true));
        this.Columns["Id"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.DG_QSTLANGL", tblName);
        dtm.ColumnMappings.Add("ID_ITEM", "Id");
        dtm.ColumnMappings.Add("STR_SQL", "StrSql");
        dtm.ColumnMappings.Add("STR_DLG", "StrDlg");
        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText = "SELECT ID_ITEM, STR_SQL, STR_DLG FROM VIZ_PRN.DG_QSTLANGL WHERE ID_LIST = :IDLST ORDER BY 1",
          CommandType = CommandType.Text
        };

        var param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "IDLST",
          SourceColumn = "ID_LIST",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);
      }

      public int LoadData(int typeList)
      {
        var lstPrmValue = new List<Object> {typeList};
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }

    }

    public sealed class ThicknessDataTable : DataTable
    {
      public ThicknessDataTable(string tblName) : base()
      {
        this.TableName = tblName;
        var col = new DataColumn("Id", typeof(decimal), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("StrDlg", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { this.Columns["Id"] }, true));
        this.Columns["Id"].Unique = true;

        System.Data.DataRow row = this.NewRow();
        row[0] = 0.23;
        row[1] = "0,23";
        this.Rows.Add(row);

        row = this.NewRow();
        row[0] = 0.27;
        row[1] = "0,27";
        this.Rows.Add(row);

        row = this.NewRow();
        row[0] = 0.30;
        row[1] = "0,30";
        this.Rows.Add(row);

        row = this.NewRow();
        row[0] = 0.35;
        row[1] = "0,35";
        this.Rows.Add(row);

        this.AcceptChanges();
      }
 
    }


  }
}
