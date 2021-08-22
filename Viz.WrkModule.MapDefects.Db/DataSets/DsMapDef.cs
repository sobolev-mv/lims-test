using System;
using System.Data;
using System.Collections.Generic;
using Smv.Data.Oracle;
using Devart.Data.Oracle;

namespace Viz.WrkModule.MapDefects.Db.DataSets
{
  public sealed class DsMapDef : DataSet
  {
    public MapDefDataTable MapDef {get; private set;}
    public LstDefZonesDataTable LstDefZones { get; private set; }
    public CutMatDataTable CutMat { get; private set; }

    public DsMapDef() : base()
    {
      this.DataSetName = "DsMapDef";

      this.MapDef = new MapDefDataTable();
      this.Tables.Add(this.MapDef);

      this.LstDefZones = new LstDefZonesDataTable();
      this.Tables.Add(this.LstDefZones);

      this.CutMat = new CutMatDataTable();
      this.Tables.Add(this.CutMat);
    }

    public sealed class MapDefDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public MapDefDataTable() : base()
      {
        this.TableName = "MapDef";
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Rid", typeof(Int64), null, MappingType.Element) {AllowDBNull = false};
        this.Columns.Add(col);

        col = new DataColumn("Zdn", typeof(Int64), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("CoilNo", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("FehlerTyp", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("WeightFrom", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("WeightTo", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("XposvOn", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("XposbIs", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("YposvOn", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("YposbIs", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("DefectSide", typeof(Int32), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("LfdNr", typeof(Int32), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Cat", typeof(String), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("AusPraeg", typeof(String), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("ZoneFrom", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("ZoneTo", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Ylen", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        this.Constraints.Add(new UniqueConstraint("Pk_MapDef", new[] { this.Columns["Rid"] }, true));
        this.Columns["Rid"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.OTK_DEF", "MapDef");
        dtm.ColumnMappings.Add("RID", "Rid");
        dtm.ColumnMappings.Add("ZDN", "Zdn");
        dtm.ColumnMappings.Add("COILNO", "CoilNo");
        dtm.ColumnMappings.Add("FEHLERTYP", "FehlerTyp");
        dtm.ColumnMappings.Add("WEIGHTFROM", "WeightFrom");
        dtm.ColumnMappings.Add("WEIGHTTO", "WeightTo");
        dtm.ColumnMappings.Add("XPOSVON", "XposvOn");
        dtm.ColumnMappings.Add("XPOSBIS", "XposbIs");
        dtm.ColumnMappings.Add("YPOSVON", "YposvOn");
        dtm.ColumnMappings.Add("YPOSBIS", "YposbIs");
        dtm.ColumnMappings.Add("DEFECT_SIDE", "DefectSide");
        dtm.ColumnMappings.Add("LFD_NR", "LfdNr");
        dtm.ColumnMappings.Add("CAT", "Cat");
        dtm.ColumnMappings.Add("AUSPRAEGUNG", "AusPraeg");
        dtm.ColumnMappings.Add("ZONEFROM", "ZoneFrom");
        dtm.ColumnMappings.Add("ZONETO", "ZoneTo");
        dtm.ColumnMappings.Add("YLEN", "Ylen");
        adapter.TableMappings.Add(dtm);
        adapter.SelectCommand = new OracleCommand
                                {
                                  Connection = Odac.DbConnection,
                                  /*
                                  CommandText = "SELECT ROWNUM RID, ZDN, COILNO, FEHLERTYP, WEIGHTFROM, WEIGHTTO, XPOSVON, XPOSBIS, YPOSVON, YPOSBIS, DEFECT_SIDE, LFD_NR, CAT, AUSPRAEGUNG, ZONEFROM, ZONETO, YLEN " +
                                                "FROM ( " +
                                                "SELECT ZDN, COILNO, FEHLERTYP, WEIGHTFROM, WEIGHTTO, XPOSVON, XPOSBIS, YPOSVON, YPOSBIS, DEFECT_SIDE, LFD_NR, CAT, AUSPRAEGUNG, ZONEFROM, ZONETO, YLEN " +
                                                "FROM VIZ_PRN.OTK_DEF WHERE (ZDN = :PZDN) AND DEFECT_SIDE IN (:DS1, :DS2) " +
                                                "UNION ALL " +
                                                "SELECT ZDN, COILNO, FEHLERTYP, WEIGHTFROM, WEIGHTTO, XPOSVON, XPOSBIS, YPOSVON, YPOSBIS, 3 AS DEFECT_SIDE, LFD_NR, CAT, AUSPRAEGUNG, ZONEFROM, ZONETO, YLEN " +
                                                "FROM VIZ_PRN.OTK_DEF WHERE (ZDN = :PZDN) AND (FEHLERTYP = 'WELDSEAM') " +
                                                ") " +
                                                "ORDER BY ZONEFROM, ZONETO, TO_NUMBER(AUSPRAEGUNG) DESC, YLEN DESC NULLS LAST",
                                  */
                                  CommandType = CommandType.Text
                                };
        
      }

      public int LoadData(long Zdn, int Side1, int Side2)
      {
        adapter.SelectCommand.Parameters.Clear();

        adapter.SelectCommand.CommandText =
          "SELECT ROWNUM RID, ZDN, COILNO, FEHLERTYP, WEIGHTFROM, WEIGHTTO, XPOSVON, XPOSBIS, YPOSVON, YPOSBIS, DEFECT_SIDE, LFD_NR, CAT, AUSPRAEGUNG, ZONEFROM, ZONETO, YLEN " +
          "FROM VIZ_PRN.OTK_DEF WHERE (ZDN = :PZDN) AND DEFECT_SIDE IN (:DS1, :DS2) " +
          "ORDER BY ZONEFROM, ZONETO, TO_NUMBER(AUSPRAEGUNG) DESC, YLEN DESC NULLS LAST";

        var param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "PZDN",
          SourceColumn = "ZDN",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "DS1",
          SourceColumn = "DEFECT_SIDE",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "DS2",
          SourceColumn = "DEFECT_SIDE",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        var lstPrmValue = new List<Object> { Zdn, Side1, Side2 };
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }

      public int LoadDataPack(long Zdn, int Side1, int Side2)
      {
        adapter.SelectCommand.Parameters.Clear();

        adapter.SelectCommand.CommandText =
          "SELECT ROWNUM RID, ZDN, COILNO, FEHLERTYP, WEIGHTFROM, WEIGHTTO, XPOSVON, XPOSBIS, YPOSVON, YPOSBIS, DEFECT_SIDE, LFD_NR, CAT, AUSPRAEGUNG, ZONEFROM, ZONETO, YLEN " +
          "FROM ( " +
          "SELECT ZDN, COILNO, FEHLERTYP, WEIGHTFROM, WEIGHTTO, XPOSVON, XPOSBIS, YPOSVON, YPOSBIS, DEFECT_SIDE, LFD_NR, CAT, AUSPRAEGUNG, ZONEFROM, ZONETO, YLEN " +
          "FROM VIZ_PRN.OTK_DEF WHERE (ZDN = :PZDN) AND DEFECT_SIDE IN (:DS1, :DS2) " +
          "UNION ALL " +
          "SELECT ZDN, COILNO, FEHLERTYP, WEIGHTFROM, WEIGHTTO, XPOSVON, XPOSBIS, YPOSVON, YPOSBIS, 3 AS DEFECT_SIDE, LFD_NR, CAT, AUSPRAEGUNG, ZONEFROM, ZONETO, YLEN " +
          "FROM VIZ_PRN.OTK_DEF WHERE (ZDN = :PZDN) AND (FEHLERTYP = 'WELDSEAM') " +
          ") " +
          "ORDER BY ZONEFROM, ZONETO, TO_NUMBER(AUSPRAEGUNG) DESC, YLEN DESC NULLS LAST";

        var param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "PZDN",
          SourceColumn = "ZDN",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "DS1",
          SourceColumn = "DEFECT_SIDE",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "DS2",
          SourceColumn = "DEFECT_SIDE",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        var lstPrmValue = new List<Object> { Zdn, Side1, Side2 };
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }
      
    }

    public sealed class LstDefZonesDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public LstDefZonesDataTable() : base()
      {
        this.TableName = "LstDefZones";
        adapter = new OracleDataAdapter();

        var col = new DataColumn("ZoneFrom", typeof(decimal), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("ZoneTo", typeof(decimal), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("Cnt", typeof(int), null, MappingType.Element);
        this.Columns.Add(col);

        this.Constraints.Add(new UniqueConstraint("Pk_LstDefZones", new[] { this.Columns["ZoneFrom"], this.Columns["ZoneTo"] }, true));
        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.T999", "LstDefZones");
        dtm.ColumnMappings.Add("ZONEFROM", "ZoneFrom");
        dtm.ColumnMappings.Add("ZONETO", "ZoneTo");
        dtm.ColumnMappings.Add("CNT", "Cnt");
        adapter.TableMappings.Add(dtm);
        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText = "SELECT ZONEFROM, ZONETO, COUNT(*) CNT FROM VIZ_PRN.OTK_DEF WHERE (ZDN = :ZDN) AND DEFECT_SIDE IN (:DS1, :DS2) GROUP BY ZONEFROM, ZONETO ORDER BY 1,2",
          CommandType = CommandType.Text
        };

        var param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "ZDN",
          SourceColumn = "ZDN",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "DS1",
          SourceColumn = "DEFECT_SIDE",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "DS2",
          SourceColumn = "DEFECT_SIDE",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);
      }

      public int LoadData(Int64 Zdn, int Side1, int Side2)
      {
        var lstPrmValue = new List<Object> {Zdn, Side1, Side2};
        int rez = Odac.LoadDataTable(this, adapter, true, lstPrmValue);

        /*
        var lstPrm = new List<OracleParameter>();
        var prm = new OracleParameter()
        {
          ParameterName = "PZDN",
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          Value = Zdn
        };
        lstPrm.Add(prm);
        Odac.ExecuteNonQuery("DELETE FROM VIZ_PRN.OTK_DEF WHERE (ZDN = :PZDN)", CommandType.Text, false, lstPrm);
        */
        return rez;
      }
    }

    public sealed class CutMatDataTable : DataTable
    {
      private readonly OracleDataAdapter adapter;

      public CutMatDataTable() : base()
      {
        this.TableName = "CutMat";
        adapter = new OracleDataAdapter();

        var col = new DataColumn("MatChild", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("XstartAncWgt", typeof(double), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("XendAncWgt", typeof(double), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("YstartAnc", typeof(int), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("YendAnc", typeof(int), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("WeightAnc", typeof(double), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Weight", typeof(double), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Sort", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Cat", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Def", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Status", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Xpart", typeof(double), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("YstartChaild", typeof(int), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("YendChaild", typeof(int), null, MappingType.Element);
        this.Columns.Add(col);


        this.Constraints.Add(new UniqueConstraint("Pk_CutMat", new[] { this.Columns["MatChild"]}, true));
        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.CUTMATG_AVO", "CutMat");
        dtm.ColumnMappings.Add("MAT_CHILD", "MatChild");
        dtm.ColumnMappings.Add("XSTARTANC_WGT", "XstartAncWgt");
        dtm.ColumnMappings.Add("XENDANC_WGT", "XendAncWgt");
        dtm.ColumnMappings.Add("YSTARTANC", "YstartAnc");
        dtm.ColumnMappings.Add("YENDANC", "YendAnc");
        dtm.ColumnMappings.Add("WEIGHTANC", "WeightAnc");
        dtm.ColumnMappings.Add("WEIGHT", "Weight");
        dtm.ColumnMappings.Add("SORT", "Sort");
        dtm.ColumnMappings.Add("CAT", "Cat");
        dtm.ColumnMappings.Add("DEF", "Def");
        dtm.ColumnMappings.Add("STATUS", "Status");
        dtm.ColumnMappings.Add("YSTARTCHILD", "YstartChaild");
        dtm.ColumnMappings.Add("YENDCHILD", "YendChaild");
        adapter.TableMappings.Add(dtm);
        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText = "select MAT_CHILD, XSTARTANC_WGT, XENDANC_WGT, YSTARTANC, YENDANC, WEIGHTANC, WEIGHT, SORT, CAT, DEF, STATUS, XPART, YSTARTCHILD, YENDCHILD " +
                        "from VIZ_PRN.CUTMATG_AVO",
          CommandType = CommandType.Text
        };
      }

      public int LoadData()
      {
       int rez = Odac.LoadDataTable(this, adapter, true, null);
       return rez;
      }



    }





  }
}
