using System;
using System.Data;
using System.Collections.Generic;
using Smv.Data.Oracle;
using Devart.Data.Oracle;
using Smv.Data;

namespace Viz.WrkModule.Isc.Db.DataSets
{
  public sealed class DsIsc : DataSet
  {

    public ShipProdPropDataTable ShipProdProp { get; private set; }
    public AgregateDataTable Agregate { get; private set; }
    public ShiftDataTable Shift { get; private set; }
    public ProductDataTable Product { get; private set; }
    public DtRespDataTable DtResp { get; private set; }
    public DownTimeDataTable DownTime { get; private set; }
    public MnfDataTable Mnf { get; private set; }

    public DsIsc() : base()
    {
      this.DataSetName = "DsIsc";

      this.ShipProdProp = new ShipProdPropDataTable("ShipProdProp");
      this.Tables.Add(this.ShipProdProp);

      this.Agregate = new AgregateDataTable("Agregate");
      this.Tables.Add(this.Agregate);

      this.Shift = new ShiftDataTable("Shift");
      this.Tables.Add(this.Shift);

      this.Product = new ProductDataTable("Product");
      this.Tables.Add(this.Product);

      this.DtResp = new DtRespDataTable("DtResp");
      this.Tables.Add(this.DtResp);

      this.DownTime = new DownTimeDataTable("DownTime");
      this.Tables.Add(this.DownTime);

      this.Mnf = new MnfDataTable("Mnf");
      this.Tables.Add(this.Mnf);

    }

    public sealed class ShipProdPropDataTable : SmvDataTable
    {
      private readonly OracleDataAdapter adapter;

      public ShipProdPropDataTable(string tblName) : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("MeId", typeof(string), null, MappingType.Element) /*{ AllowDBNull = false }*/;
        this.Columns.Add(col);

        col = new DataColumn("AuNr", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("UposNr", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("LocNo", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("DateShipping", typeof(DateTime), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("ContractNo", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("SpecNo", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Net", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Gross", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Thickness", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Width", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("P1550Ap", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("P1750Ap", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("P1750Lst", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("B800Lst", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("B800Ap", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("NumOfWelds", typeof(Int32), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("HeatNo", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("StoGrade", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("KesiAvg", typeof(Int32), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Gib", typeof(Int32), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("PlacementNum", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("AnnealingLot", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Grade", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Standart", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("SertNo", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("NameMnf", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("TypeMnf", typeof(Int32), null, MappingType.Element);
        this.Columns.Add(col);


        //this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { this.Columns["MeId"] }, true));
        //this.Columns["MeId"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.V_ISC_SPP_ALL", tblName);
        dtm.ColumnMappings.Add("ME_ID", "MeId");
        dtm.ColumnMappings.Add("AU_NR", "AuNr");
        dtm.ColumnMappings.Add("UPOS_NR", "UposNr");
        dtm.ColumnMappings.Add("LOCNO", "LocNo");
        dtm.ColumnMappings.Add("DATE_SHIPPING", "DateShipping");
        dtm.ColumnMappings.Add("CONTRACTNO", "ContractNo");
        dtm.ColumnMappings.Add("SPECNO", "SpecNo");
        dtm.ColumnMappings.Add("NET", "Net");
        dtm.ColumnMappings.Add("GROSS", "Gross");
        dtm.ColumnMappings.Add("THICKNESS", "Thickness");
        dtm.ColumnMappings.Add("WIDTH", "Width");
        dtm.ColumnMappings.Add("P1550AP", "P1550Ap");
        dtm.ColumnMappings.Add("P1750AP", "P1750Ap");
        dtm.ColumnMappings.Add("P1750LST", "P1750Lst");
        dtm.ColumnMappings.Add("B800LST", "B800Lst");
        dtm.ColumnMappings.Add("B800AP", "B800Ap");
        dtm.ColumnMappings.Add("NUMOFWELDS", "NumOfWelds");
        dtm.ColumnMappings.Add("HEATNO", "HeatNo");
        dtm.ColumnMappings.Add("STOGRADE", "StoGrade");
        dtm.ColumnMappings.Add("KESIAVG", "KesiAvg");
        dtm.ColumnMappings.Add("GIB", "Gib");
        dtm.ColumnMappings.Add("PLACEMENT_NUM", "PlacementNum");
        dtm.ColumnMappings.Add("ANNEALINGLOT", "AnnealingLot");
        dtm.ColumnMappings.Add("GRADE", "Grade");
        dtm.ColumnMappings.Add("STANDART", "Standart");
        dtm.ColumnMappings.Add("CERT_NR", "SertNo");
        dtm.ColumnMappings.Add("NAME_MNF", "NameMnf");
        dtm.ColumnMappings.Add("TYP_MNF", "TypeMnf");

        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText = "SELECT " +
                        "ME_ID, AU_NR, UPOS_NR, LOCNO, DATE_SHIPPING, CONTRACTNO, SPECNO, NET, GROSS, THICKNESS, WIDTH, P1550AP, P1750AP, P1750LST, B800LST, " +
                        "B800AP, NUMOFWELDS, HEATNO, STOGRADE, KESIAVG, GIB, PLACEMENT_NUM, ANNEALINGLOT, GRADE, STANDART, CERT_NR, NAME_MNF, TYP_MNF " +
                        "FROM VIZ_PRN.V_ISC_SPP_ALL " +
                        "WHERE ((DATE_SHIPPING BETWEEN :DSH1 AND :DSH2) OR (:ISDSH = 0)) " +
                        "AND ((CONTRACTNO = :CNTNO) OR (:ISCNTNO = 0)) " +
                        "AND ((SPECNO = :SPCNO) OR (:ISSPCNO = 0)) " +
                        "AND ((CERT_NR = :CERTNR) OR (:ISCERTNR = 0)) " +
                        "AND ((TYP_MNF = :TYPMNF) OR (:ISTYPMNF = 0)) " +
                        "AND ((PLACEMENT_NUM = :PLACEMENTNUM) OR (:ISPLACEMENTNUM = 0))",
          
          CommandType = CommandType.Text
        };
        
        var param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "DSH1",
          SourceColumn = "DATE_SHIPPING",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "DSH2",
          SourceColumn = "DATE_SHIPPING",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISDSH",
          SourceColumn = "DATE_SHIPPING",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "CNTNO",
          SourceColumn = "CONTRACTNO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISCNTNO",
          SourceColumn = "CONTRACTNO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);
        
        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "SPCNO",
          SourceColumn = "SPECNO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISSPCNO",
          SourceColumn = "SPECNO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "CERTNR",
          SourceColumn = "CERT_NR",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISCERTNR",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "TYPMNF",
          SourceColumn = "TYP_MNF",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISTYPMNF",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PLACEMENTNUM",
          SourceColumn = "PLACEMENT_NUM",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "ISPLACEMENTNUM",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);


      }

      public int LoadData(DateTime dtShip1, DateTime dtShip2, Boolean isDtShip, string contractNo, Boolean isContractNo, string specNo, Boolean isSpecNo, string sertNo, Boolean isSertNo, Int32 typMnf, Boolean isTypMnf, string placementNo, Boolean isPlacementNo)
      {
        var lstPrmValue = new List<Object> { dtShip1, dtShip2, isDtShip ? 1 : 0, contractNo, isContractNo ? 1 : 0, specNo, isSpecNo ? 1 : 0, sertNo, isSertNo ? 1 : 0, typMnf, isTypMnf ? 1 : 0, placementNo, isPlacementNo ? 1 : 0};
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }

    }

    public sealed class MnfDataTable : SmvDataTable
    {
      private readonly OracleDataAdapter adapter;

      public MnfDataTable(string tblName) : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof(Int32), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("NameMnf", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);


        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { this.Columns["Id"] }, true));
        this.Columns["Id"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.ISC_MNF", tblName);
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("NAME_MNF", "NameMnf");
        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText = "SELECT ID,  NAME_MNF " +
                        "FROM VIZ_PRN.ISC_MNF " +
                        "ORDER BY 1",
          CommandType = CommandType.Text
        };

      }

      public int LoadData()
      {
        return Odac.LoadDataTable(this, adapter, true, null);
      }

    }




    public sealed class AgregateDataTable : SmvDataTable
    {
      private readonly OracleDataAdapter adapter;

      public AgregateDataTable(string tblName)
        : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("NameAgr", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);


        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { this.Columns["Id"] }, true));
        this.Columns["Id"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.ISC_AGR", tblName);
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("NAME_AGR", "NameAgr");
        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText = "SELECT ID, NAME_AGR " +
                        "FROM VIZ_PRN.ISC_AGR " +
                        "WHERE IS_ACTIVE = 'Y' ORDER BY 1",
          CommandType = CommandType.Text
        };
        
      }

      public int LoadData()
      {
        return Odac.LoadDataTable(this, adapter, true, null);
      }

    }

    public class ShiftDataTable : SmvDataTable
    {
      protected readonly OracleDataAdapter adapter;

      public ShiftDataTable(string tblName)
        : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof(Int64), null, MappingType.Element)
        {
          AllowDBNull = false,
          AutoIncrement = true,
          AutoIncrementStep = -1
        };
        this.Columns.Add(col);

        col = new DataColumn("DateShift", typeof(DateTime), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("Shift", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("AgrId", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("Team", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("TeamMembers", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);
          

        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { this.Columns["Id"] }, true));
        this.Columns["Id"].Unique = true;
        this.Constraints.Add(new UniqueConstraint("Pk2_" + tblName, new[] { this.Columns["DateShift"], this.Columns["Shift"], this.Columns["AgrId"] }, false));

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.ISC_SHIFT", tblName);
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("DATE_SHIFT", "DateShift");
        dtm.ColumnMappings.Add("SHIFT", "Shift");
        dtm.ColumnMappings.Add("AGR_ID", "AgrId");
        dtm.ColumnMappings.Add("TEAM", "Team");
        dtm.ColumnMappings.Add("TEAM_MEMBERS", "TeamMembers");

        adapter.TableMappings.Add(dtm);

        //Select Command
        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "SELECT ID, DATE_SHIFT, SHIFT, AGR_ID, TEAM, TEAM_MEMBERS " +
            "FROM VIZ_PRN.ISC_SHIFT WHERE (DATE_SHIFT BETWEEN :DT1 AND :DT2) AND (AGR_ID = :AGRID) ORDER BY DATE_SHIFT",
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

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "AGRID",
          SourceColumn = "AGR_ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);


        //Insert Command
        adapter.InsertCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "INSERT INTO VIZ_PRN.ISC_SHIFT(DATE_SHIFT, SHIFT, AGR_ID, TEAM, TEAM_MEMBERS) " +
            "VALUES(:PDATE_SHIFT, :PSHIFT, :PAGR_ID, :PTEAM, :PTEAM_MEMBERS) RETURNING ID INTO :PID",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.OutputParameters
        };

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_SHIFT",
          SourceColumn = "DATE_SHIFT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PSHIFT",
          SourceColumn = "SHIFT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PAGR_ID",
          SourceColumn = "AGR_ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTEAM",
          SourceColumn = "TEAM",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTEAM_MEMBERS",
          SourceColumn = "TEAM_MEMBERS",
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

        //Update Command
        adapter.UpdateCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "UPDATE VIZ_PRN.ISC_SHIFT SET DATE_SHIFT = :PDATE_SHIFT, SHIFT = :PSHIFT, TEAM = :PTEAM, TEAM_MEMBERS = :PTEAM_MEMBERS " +
            "WHERE (ID = :Original_ID)",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.None
        };

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_SHIFT",
          SourceColumn = "DATE_SHIFT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PSHIFT",
          SourceColumn = "SHIFT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTEAM",
          SourceColumn = "TEAM",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTEAM_MEMBERS",
          SourceColumn = "TEAM_MEMBERS",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);
        
        param = new OracleParameter
        {
          DbType = DbType.Int64,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "Original_ID",
          SourceColumn = "ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Original
        };
        adapter.UpdateCommand.Parameters.Add(param);

        //Delete Command
        adapter.DeleteCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "DELETE FROM VIZ_PRN.ISC_SHIFT " +
            "WHERE (ID = :Original_ID)",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.None
        };

        param = new OracleParameter
        {
          DbType = DbType.Int64,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "Original_ID",
          SourceColumn = "ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Original
        };
        adapter.DeleteCommand.Parameters.Add(param);

      }

      public int LoadData(DateTime dtBegin, DateTime dtEnd, string agrId)
      {
        var lstPrmValue = new List<Object> { dtBegin, dtEnd, agrId };
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }


      public int SaveData()
      {
        return Odac.SaveChangedData(this, adapter);
      }

    }

    public class ProductDataTable : SmvDataTable
    {
      protected readonly OracleDataAdapter adapter;

      public ProductDataTable(string tblName)
        : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof(Int64), null, MappingType.Element)
        {
          AllowDBNull = false,
          AutoIncrement = true,
          AutoIncrementStep = -1
        };
        this.Columns.Add(col);

        col = new DataColumn("ShiftId", typeof(Int64), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("CoilNo", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("LotNo", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("Weight", typeof(decimal), null, MappingType.Element) { DefaultValue = 0};
        this.Columns.Add(col);

        col = new DataColumn("CoilNoNext", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Thickness", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Width", typeof(Int32), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Sort", typeof(string), null, MappingType.Element) { DefaultValue = 1 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp1", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp2", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp3", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp4", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp5", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp6", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp7", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp8", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp9", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp10", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp11", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Ysp12", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("YeldWeight", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("EdgeCrop", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("CrossCut", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Residues", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("WeldJoin", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("Choice", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("NameItem", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("CoilLength", typeof(decimal), null, MappingType.Element) { DefaultValue = 0 };
        this.Columns.Add(col);

        col = new DataColumn("TxtComment", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);


        
        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { this.Columns["Id"] }, true));
        this.Columns["Id"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("LIMS.THP_DATA", tblName);
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("SHIFT_ID", "ShiftId");
        dtm.ColumnMappings.Add("COIL_NO", "CoilNo");
        dtm.ColumnMappings.Add("LOT_NO", "LotNo");
        dtm.ColumnMappings.Add("WEIGHT", "Weight");
        dtm.ColumnMappings.Add("COIL_NO_NEXT", "CoilNoNext");
        dtm.ColumnMappings.Add("THICKNESS", "Thickness");
        dtm.ColumnMappings.Add("WIDTH", "Width");
        dtm.ColumnMappings.Add("SORT", "Sort");
        dtm.ColumnMappings.Add("YSP1", "Ysp1");
        dtm.ColumnMappings.Add("YSP2", "Ysp2");
        dtm.ColumnMappings.Add("YSP3", "Ysp3");
        dtm.ColumnMappings.Add("YSP4", "Ysp4");
        dtm.ColumnMappings.Add("YSP5", "Ysp5");
        dtm.ColumnMappings.Add("YSP6", "Ysp6");
        dtm.ColumnMappings.Add("YSP7", "Ysp7");
        dtm.ColumnMappings.Add("YSP8", "Ysp8");
        dtm.ColumnMappings.Add("YSP9", "Ysp9");
        dtm.ColumnMappings.Add("YSP10", "Ysp10");
        dtm.ColumnMappings.Add("YSP11", "Ysp11");
        dtm.ColumnMappings.Add("YSP12", "Ysp12");
        dtm.ColumnMappings.Add("YELD_WEIGHT", "YeldWeight");
        dtm.ColumnMappings.Add("EDGE_CROP", "EdgeCrop");
        dtm.ColumnMappings.Add("CROSS_CUT", "CrossCut");
        dtm.ColumnMappings.Add("RESIDUES", "Residues");
        dtm.ColumnMappings.Add("WELD_JOIN", "WeldJoin");
        dtm.ColumnMappings.Add("CHOICE", "Choice");
        dtm.ColumnMappings.Add("NAME_ITEM", "NameItem");
        dtm.ColumnMappings.Add("COIL_LENGTH", "CoilLength");
        dtm.ColumnMappings.Add("TXTCOMMENT", "TxtComment");
        adapter.TableMappings.Add(dtm);

        //Select Command
        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "SELECT ID,SHIFT_ID,COIL_NO,LOT_NO,WEIGHT,COIL_NO_NEXT,THICKNESS,WIDTH,SORT,YSP1,YSP2,YSP3,YSP4,YSP5,YSP6,YSP7,YSP8,YSP9,YSP10,YSP11, " +
            "YSP12,YELD_WEIGHT,EDGE_CROP,CROSS_CUT,RESIDUES,WELD_JOIN,CHOICE,NAME_ITEM,COIL_LENGTH, TXTCOMMENT " +
            "FROM VIZ_PRN.ISC_PDODUCT WHERE SHIFT_ID = :PSHIFT_ID ORDER BY ID",
          CommandType = CommandType.Text
        };

        var param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PSHIFT_ID",
          SourceColumn = "SHIFT_ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 15,
          Scale = 0
        };
        adapter.SelectCommand.Parameters.Add(param);

        //Insert Command
        adapter.InsertCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "INSERT INTO VIZ_PRN.ISC_PDODUCT(SHIFT_ID,COIL_NO,LOT_NO,WEIGHT,COIL_NO_NEXT,THICKNESS,WIDTH,SORT,YSP1,YSP2,YSP3,YSP4,YSP5,YSP6,YSP7,YSP8,YSP9,YSP10,YSP11, " +
            "YSP12,YELD_WEIGHT,EDGE_CROP,CROSS_CUT,RESIDUES,WELD_JOIN,CHOICE,NAME_ITEM,COIL_LENGTH, TXTCOMMENT) " +
            "VALUES(:PSHIFT_ID, :PCOIL_NO, :PLOT_NO, :PWEIGHT, :PCOIL_NO_NEXT, :PTHICKNESS, :PWIDTH, :PSORT, :PYSP1, :PYSP2, :PYSP3, :PYSP4, :PYSP5, :PYSP6, :PYSP7, :PYSP8, :PYSP9, :PYSP10, :PYSP11, " +
            ":PYSP12, :PYELD_WEIGHT, :PEDGE_CROP, :PCROSS_CUT, :PRESIDUES, :PWELD_JOIN, :PCHOICE, :PNAME_ITEM, :PCOIL_LENGTH, :PTXTCOMMENT) RETURNING ID INTO :PID",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.OutputParameters
        };

        param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PSHIFT_ID",
          SourceColumn = "SHIFT_ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PCOIL_NO",
          SourceColumn = "COIL_NO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PLOT_NO",
          SourceColumn = "LOT_NO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PWEIGHT",
          SourceColumn = "WEIGHT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PCOIL_NO_NEXT",
          SourceColumn = "COIL_NO_NEXT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PTHICKNESS",
          SourceColumn = "THICKNESS",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 5,
          Scale = 2
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "PWIDTH",
          SourceColumn = "WIDTH",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);
        
        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PSORT",
          SourceColumn = "SORT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP1",
          SourceColumn = "YSP1",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP2",
          SourceColumn = "YSP2",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP3",
          SourceColumn = "YSP3",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP4",
          SourceColumn = "YSP4",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP5",
          SourceColumn = "YSP5",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP6",
          SourceColumn = "YSP6",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP7",
          SourceColumn = "YSP7",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP8",
          SourceColumn = "YSP8",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP9",
          SourceColumn = "YSP9",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP10",
          SourceColumn = "YSP10",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP11",
          SourceColumn = "YSP11",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP12",
          SourceColumn = "YSP12",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYELD_WEIGHT",
          SourceColumn = "YELD_WEIGHT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PEDGE_CROP",
          SourceColumn = "EDGE_CROP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PCROSS_CUT",
          SourceColumn = "CROSS_CUT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PRESIDUES",
          SourceColumn = "RESIDUES",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PWELD_JOIN",
          SourceColumn = "WELD_JOIN",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PCHOICE",
          SourceColumn = "CHOICE",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PNAME_ITEM",
          SourceColumn = "NAME_ITEM",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PCOIL_LENGTH",
          SourceColumn = "COIL_LENGTH",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTXTCOMMENT",
          SourceColumn = "TXTCOMMENT",
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


        //Update Command
        adapter.UpdateCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "UPDATE VIZ_PRN.ISC_PDODUCT SET " +
            "COIL_NO = :PCOIL_NO, LOT_NO = :PLOT_NO, WEIGHT = :PWEIGHT, COIL_NO_NEXT = :PCOIL_NO_NEXT, THICKNESS = :PTHICKNESS, WIDTH = :PWIDTH, SORT = :PSORT, " +
            "YSP1 = :PYSP1, YSP2 = :PYSP2, YSP3 = :PYSP3, YSP4 = :PYSP4, YSP5 = :PYSP5, YSP6 = :PYSP6, YSP7 = :PYSP7, YSP8 = :PYSP8, YSP9 = :PYSP9, YSP10 = :PYSP10, YSP11 = :PYSP11, " +
            "YSP12 = :PYSP12, YELD_WEIGHT = :PYELD_WEIGHT, EDGE_CROP = :PEDGE_CROP, CROSS_CUT = :PCROSS_CUT, RESIDUES = :PRESIDUES, WELD_JOIN = :PWELD_JOIN, CHOICE = :PCHOICE, " +
            "NAME_ITEM = :PNAME_ITEM, COIL_LENGTH = :PCOIL_LENGTH, TXTCOMMENT = :PTXTCOMMENT " +
            "WHERE (ID = :Original_ID)",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.None
        };

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PCOIL_NO",
          SourceColumn = "COIL_NO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PLOT_NO",
          SourceColumn = "LOT_NO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PWEIGHT",
          SourceColumn = "WEIGHT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PCOIL_NO_NEXT",
          SourceColumn = "COIL_NO_NEXT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PTHICKNESS",
          SourceColumn = "THICKNESS",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 5,
          Scale = 2
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int32,
          OracleDbType = OracleDbType.Integer,
          Direction = ParameterDirection.Input,
          ParameterName = "PWIDTH",
          SourceColumn = "WIDTH",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PSORT",
          SourceColumn = "SORT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP1",
          SourceColumn = "YSP1",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP2",
          SourceColumn = "YSP2",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP3",
          SourceColumn = "YSP3",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP4",
          SourceColumn = "YSP4",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP5",
          SourceColumn = "YSP5",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP6",
          SourceColumn = "YSP6",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP7",
          SourceColumn = "YSP7",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP8",
          SourceColumn = "YSP8",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP9",
          SourceColumn = "YSP9",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP10",
          SourceColumn = "YSP10",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP11",
          SourceColumn = "YSP11",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYSP12",
          SourceColumn = "YSP12",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 10,
          Scale = 1
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PYELD_WEIGHT",
          SourceColumn = "YELD_WEIGHT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PEDGE_CROP",
          SourceColumn = "EDGE_CROP",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PCROSS_CUT",
          SourceColumn = "CROSS_CUT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PRESIDUES",
          SourceColumn = "RESIDUES",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PWELD_JOIN",
          SourceColumn = "WELD_JOIN",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PCHOICE",
          SourceColumn = "CHOICE",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PNAME_ITEM",
          SourceColumn = "NAME_ITEM",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PCOIL_LENGTH",
          SourceColumn = "COIL_LENGTH",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current,
          Precision = 17,
          Scale = 3
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTXTCOMMENT",
          SourceColumn = "TXTCOMMENT",
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

        //Delete Command
        adapter.DeleteCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "DELETE FROM VIZ_PRN.ISC_PDODUCT " +
            "WHERE (ID = :Original_ID)",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.None
        };

        param = new OracleParameter
        {
          DbType = DbType.Int64,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "Original_ID",
          SourceColumn = "ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Original
        };
        adapter.DeleteCommand.Parameters.Add(param);

      }

      public int LoadData(Int64 shiftId)
      {
        var lstPrmValue = new List<Object> { shiftId };
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }

      public int SaveData()
      {
        return Odac.SaveChangedData(this, adapter);
      }

    }

    public sealed class DtRespDataTable : SmvDataTable
    {
      private readonly OracleDataAdapter adapter;

      public DtRespDataTable(string tblName)
        : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("NameResp", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);


        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { this.Columns["Id"] }, true));
        this.Columns["Id"].Unique = true;

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.ISC_DTRESP", tblName);
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("NAME_RESP", "NameResp");
        adapter.TableMappings.Add(dtm);

        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText = "SELECT ID, NAME_RESP " +
                        "FROM VIZ_PRN.ISC_DTRESP " +
                        "WHERE IS_ACTIVE = 'Y' ORDER BY 1",
          CommandType = CommandType.Text
        };

      }

      public int LoadData()
      {
        return Odac.LoadDataTable(this, adapter, true, null);
      }

    }

    public class DownTimeDataTable : SmvDataTable
    {
      protected readonly OracleDataAdapter adapter;

      public DownTimeDataTable(string tblName)
        : base()
      {
        this.TableName = tblName;
        adapter = new OracleDataAdapter();

        var col = new DataColumn("Id", typeof(Int64), null, MappingType.Element)
        {
          AllowDBNull = false,
          AutoIncrement = true,
          AutoIncrementStep = -1
        };
        this.Columns.Add(col);

        col = new DataColumn("DateFrom", typeof(DateTime), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("DateTo", typeof(DateTime), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("ShiftId", typeof(Int64), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("RespId", typeof(string), null, MappingType.Element) { AllowDBNull = false };
        this.Columns.Add(col);

        col = new DataColumn("TxtComment", typeof(string), null, MappingType.Element);
        this.Columns.Add(col);

        col = new DataColumn("Duration", typeof(decimal), null, MappingType.Element);
        this.Columns.Add(col);


        this.Constraints.Add(new UniqueConstraint("Pk_" + tblName, new[] { this.Columns["Id"] }, true));
        this.Columns["Id"].Unique = true;
        

        adapter.TableMappings.Clear();
        var dtm = new System.Data.Common.DataTableMapping("VIZ_PRN.ISC_DT", tblName);
        dtm.ColumnMappings.Add("ID", "Id");
        dtm.ColumnMappings.Add("DATE_FROM", "DateFrom");
        dtm.ColumnMappings.Add("DATE_TO", "DateTo");
        dtm.ColumnMappings.Add("SHIFT_ID", "ShiftId");
        dtm.ColumnMappings.Add("RESP_ID", "RespId");
        dtm.ColumnMappings.Add("TXTCOMMENT", "TxtComment");
        dtm.ColumnMappings.Add("DURATION", "Duration");

        adapter.TableMappings.Add(dtm);

        //Select Command
        adapter.SelectCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "SELECT ID, DATE_FROM, DATE_TO, SHIFT_ID, RESP_ID, TXTCOMMENT, " +
            "VIZ_PRN.VAR_RPT.GetTub2Date(DATE_FROM, DATE_TO, 'M') DURATION " +
            "FROM VIZ_PRN.ISC_DT WHERE (SHIFT_ID = :PSHIFT_ID) ORDER BY ID",
          CommandType = CommandType.Text
        };

        var param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PSHIFT_ID",
          SourceColumn = "SHIFT_ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.SelectCommand.Parameters.Add(param);


        //Insert Command
        adapter.InsertCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "INSERT INTO VIZ_PRN.ISC_DT(DATE_FROM, DATE_TO, SHIFT_ID, RESP_ID, TXTCOMMENT) " +
            "VALUES(:PDATE_FROM, :PDATE_TO, :PSHIFT_ID, :PRESP_ID, :PTXTCOMMENT) RETURNING ID, VIZ_PRN.VAR_RPT.GetTub2Date(DATE_FROM, DATE_TO, 'M') INTO :PID, :PDURATION",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.OutputParameters
        };

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_FROM",
          SourceColumn = "DATE_FROM",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_TO",
          SourceColumn = "DATE_TO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);
        
        param = new OracleParameter
        {
          DbType = DbType.Int64,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.Input,
          ParameterName = "PSHIFT_ID",
          SourceColumn = "SHIFT_ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PRESP_ID",
          SourceColumn = "RESP_ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTXTCOMMENT",
          SourceColumn = "TXTCOMMENT",
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
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.ReturnValue,
          ParameterName = "PDURATION",
          SourceColumn = "DURATION",
          SourceColumnNullMapping = false,
          Precision = 15,
          Scale = 2,
          SourceVersion = DataRowVersion.Current
        };
        adapter.InsertCommand.Parameters.Add(param);


        //Update Command
        adapter.UpdateCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "UPDATE VIZ_PRN.ISC_DT SET DATE_FROM = :PDATE_FROM, DATE_TO = :PDATE_TO, RESP_ID = :PRESP_ID, TXTCOMMENT = :PTXTCOMMENT " +
            "WHERE (ID = :Original_ID) RETURNING VIZ_PRN.VAR_RPT.GetTub2Date(DATE_FROM, DATE_TO, 'M') INTO :PDURATION",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.OutputParameters
        };

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_FROM",
          SourceColumn = "DATE_FROM",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.DateTime,
          OracleDbType = OracleDbType.Date,
          Direction = ParameterDirection.Input,
          ParameterName = "PDATE_TO",
          SourceColumn = "DATE_TO",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);


        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PRESP_ID",
          SourceColumn = "RESP_ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.String,
          OracleDbType = OracleDbType.VarChar,
          Direction = ParameterDirection.Input,
          ParameterName = "PTXTCOMMENT",
          SourceColumn = "TXTCOMMENT",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Int64,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "Original_ID",
          SourceColumn = "ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Original
        };
        adapter.UpdateCommand.Parameters.Add(param);

        param = new OracleParameter
        {
          DbType = DbType.Decimal,
          OracleDbType = OracleDbType.Number,
          Direction = ParameterDirection.ReturnValue,
          ParameterName = "PDURATION",
          SourceColumn = "DURATION",
          SourceColumnNullMapping = false,
          Precision = 15,
          Scale = 2,
          SourceVersion = DataRowVersion.Current
        };
        adapter.UpdateCommand.Parameters.Add(param);

        //Delete Command
        adapter.DeleteCommand = new OracleCommand
        {
          Connection = Odac.DbConnection,
          CommandText =
            "DELETE FROM VIZ_PRN.ISC_DT " +
            "WHERE (ID = :Original_ID)",
          CommandType = CommandType.Text,
          PassParametersByName = true,
          UpdatedRowSource = UpdateRowSource.None
        };

        param = new OracleParameter
        {
          DbType = DbType.Int64,
          Direction = ParameterDirection.Input,
          IsNullable = false,
          ParameterName = "Original_ID",
          SourceColumn = "ID",
          SourceColumnNullMapping = false,
          SourceVersion = DataRowVersion.Original
        };
        adapter.DeleteCommand.Parameters.Add(param);

      }

      public int LoadData(Int64 shiftId)
      {
        var lstPrmValue = new List<Object> { shiftId };
        return Odac.LoadDataTable(this, adapter, true, lstPrmValue);
      }


      public int SaveData()
      {
        return Odac.SaveChangedData(this, adapter);
      }

    }




  }

}
