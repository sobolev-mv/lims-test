using System;
using System.Data;
using System.Collections.Generic;
using Smv.Data.Oracle;
using Devart.Data.Oracle;

namespace Viz.WrkModule.PrintLabel.Db.DataSets
{
  public sealed class DsPrintLabel : DataSet
  {
    public PatternDataTable LstFinishApr { get; private set; }
    public AprMatDataTable AprMat { get; private set; }
    public DsPrintLabel() : base()
    {
      this.DataSetName = "DsPrintLabel";

      this.LstFinishApr = new PatternDataTable("LstFinishApr");
      this.Tables.Add(this.LstFinishApr);

      this.AprMat = new AprMatDataTable();
      this.Tables.Add(this.AprMat);
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

    public sealed class AprMatDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public AprMatDataTable() : base()
      {
        this.TableName = "AprMat";
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Bezeichnung", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("BundleBez", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("AnnealingLot", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("AnnealingLotSeqNo", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("ErpChargenr", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Gew", typeof(int), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Dicke", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Breite", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Laenge", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("ErpMaterialnr", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Feldbez", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("AendDatum", typeof(DateTime), null, MappingType.Element);
        this.Columns.Add(col);


        this.Constraints.Add(new UniqueConstraint("Pk_" + this.TableName, new[] { this.Columns["Bezeichnung"] }, true));
        this.Columns["Bezeichnung"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.DG_QSTLANGL", this.TableName);
        dtm.ColumnMappings.Add("BEZEICHNUNG", "Bezeichnung");
        dtm.ColumnMappings.Add("BUNDLE_BEZ", "BundleBez");
        dtm.ColumnMappings.Add("ANNEALINGLOT", "AnnealingLot");
        dtm.ColumnMappings.Add("ANNEALINGLOTSEQNO", "AnnealingLotSeqNo"); 
        dtm.ColumnMappings.Add("ERPCHARGENR", "ErpChargenr");
        dtm.ColumnMappings.Add("GEW", "Gew");
        dtm.ColumnMappings.Add("DICKE", "Dicke");
        dtm.ColumnMappings.Add("BREITE", "Breite");
        dtm.ColumnMappings.Add("LAENGE", "Laenge");
        dtm.ColumnMappings.Add("ERPMATERIALNR", "ErpMaterialnr");
        dtm.ColumnMappings.Add("FELDBEZ", "Feldbez");
        dtm.ColumnMappings.Add("AENDDATUM", "AendDatum");
        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText = "SELECT BEZEICHNUNG, BUNDLE_BEZ, ANNEALINGLOT, ANNEALINGLOTSEQNO,  ERPCHARGENR, GEW, DICKE, BREITE, ROUND(LAENGE /1000,2) LAENGE, ERPMATERIALNR, FELDBEZ, AENDDATUM + 5/24 AS AENDDATUM " +
                        "FROM VIZ.D3200_MATAFTERLINE_V WHERE(ANLAGE = :PAPR) ORDER BY AENDDATUM DESC",
          CommandType = CommandType.Text
        };

        var param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PAPR",
          SourceColumn = "ANLAGE",
          SourceColumnNullMapping = true,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);
      }

      public int LoadData(string typeList)
      {
        var lstPrmValue = new List<Object> { typeList };
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }

    }





  }

}
