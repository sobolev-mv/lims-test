using System;
using System.Data;
using System.Collections.Generic;
using Smv.Data.Oracle;
using Devart.Data.Oracle;

namespace Viz.WrkModule.RptMagLab.Db.DataSets
{
  public sealed class DsRptMagLab : DataSet
  {

    public QualityTechStepDataTable Rm1200Ts   { get; private set; }
    public QualityTechStepDataTable AroTs      { get; private set; }
    public QualityTechStepDataTable AooTs      { get; private set; }
    public QualityTechStepDataTable AvoTs      { get; private set; }
    public QualityTechStepDataTable AprTs      { get; private set; }
    public QualityTechStepDataTable AdgInPrm   { get; private set; }
    public QualityTechStepDataTable AdgOutPrm  { get; private set; }
    public QualityTechStepDataTable SortTs     { get; private set; }
    public QualityTechStepDataTable ClassPloskTs { get; private set; }


    public DsRptMagLab() : base()
    {
      this.DataSetName = "DsRptMagLab";

      this.Rm1200Ts = new QualityTechStepDataTable("Rm1200Ts");
      this.Tables.Add(this.Rm1200Ts);

      this.AroTs = new QualityTechStepDataTable("AroTs");
      this.Tables.Add(this.AroTs);

      this.AooTs = new QualityTechStepDataTable("AooTs");
      this.Tables.Add(this.AooTs);

      this.AvoTs = new QualityTechStepDataTable("AvoTs");
      this.Tables.Add(this.AvoTs);

      this.AprTs = new QualityTechStepDataTable("AprTs");
      this.Tables.Add(this.AprTs);

      this.AdgInPrm = new QualityTechStepDataTable("AdgInPrm");
      this.Tables.Add(this.AdgInPrm);

      this.AdgOutPrm = new QualityTechStepDataTable("AdgOutPrm");
      this.Tables.Add(this.AdgOutPrm);

      this.SortTs = new QualityTechStepDataTable("SortTs");
      this.Tables.Add(this.SortTs);

      this.ClassPloskTs = new QualityTechStepDataTable("ClassPloskTs");
      this.Tables.Add(this.ClassPloskTs);
    }

    public sealed class QualityTechStepDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public QualityTechStepDataTable(string tblName) : base()
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

  }

}
