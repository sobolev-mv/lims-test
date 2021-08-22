using System;
using System.Data;
using System.Collections.Generic;
using Smv.Data.Oracle;
using Devart.Data.Oracle;

namespace Viz.WrkModule.RptOpr.Db.DataSets
{
  public sealed class DsRptOpr : DataSet
  {

    public PatternDataTable LstFinishApr { get; private set; }
    public PatternDataTable LstTrgtNextProc { get; private set; }
    public PatternDataTable LstTypeProd { get; private set; }
    public PatternDataTable LstThickness { get; private set; }
    public PatternDataTable LstSort { get; private set; }

    public DsRptOpr() : base()
    {
      this.DataSetName = "DsRptOpr";

      this.LstFinishApr = new PatternDataTable("LstFinishApr");
      this.Tables.Add(this.LstFinishApr);

      this.LstTrgtNextProc = new PatternDataTable("LstTrgtNextProc");
      this.Tables.Add(this.LstTrgtNextProc);

      this.LstTypeProd = new PatternDataTable("LstTypeProd");
      this.Tables.Add(this.LstTypeProd);

      this.LstThickness = new PatternDataTable("LstThickness");
      this.Tables.Add(this.LstThickness);

      this.LstSort = new PatternDataTable("LstSort");
      this.Tables.Add(this.LstSort);
    }

    public sealed class PatternDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public PatternDataTable(string tblName) : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof(Int32), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("StrSql", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("StrDlg", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { this.Columns["Id"] }, true));
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

  }

}
